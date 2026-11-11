"""Minimal Chrome DevTools Protocol client.

Deliberately avoids Runtime.enable / Page.enable: several job sites (猎聘, Boss)
detect those domains and blank the page or force extra verification.
"""
import base64
import json
import os
import subprocess
import time
import urllib.request

import websocket

JS_HELPERS = r"""
const __vis = e => { if (!e) return false; const r = e.getBoundingClientRect();
  const s = getComputedStyle(e); return r.width > 0 && r.height > 0 && s.visibility !== 'hidden' && s.display !== 'none'; };
const __byText = (re, tags) => { const rx = new RegExp(re);
  const hits = Array.from(document.querySelectorAll(tags || '*')).filter(e => { if (!__vis(e)) return false;
    const t = (e.innerText || e.value || '').trim(); return t && t.length <= 60 && rx.test(t); });
  return hits.find(e => !hits.some(o => o !== e && e.contains(o))) || null; };
const __q = (sel) => { if (sel.startsWith('text=')) return __byText(sel.slice(5));
  const all = Array.from(document.querySelectorAll(sel)); return all.find(__vis) || all[0] || null; };
"""


class CDPError(Exception):
    pass


class Browser(object):
    def __init__(self, edge_path, profile_dir, port=9333):
        self.edge_path = edge_path
        self.profile_dir = os.path.abspath(profile_dir)
        self.port = port
        self.base = 'http://127.0.0.1:%d' % port

    def _http(self, path, method='GET'):
        req = urllib.request.Request(self.base + path, method=method)
        return json.loads(urllib.request.urlopen(req, timeout=10).read().decode('utf-8'))

    def alive(self):
        try:
            self._http('/json/version')
            return True
        except Exception:
            return False

    def start(self):
        if self.alive():
            return
        os.makedirs(self.profile_dir, exist_ok=True)
        subprocess.Popen([self.edge_path, '--remote-debugging-port=%d' % self.port,
                          '--user-data-dir=' + self.profile_dir, '--no-first-run',
                          '--no-default-browser-check', '--start-maximized', 'about:blank'])
        for _ in range(60):
            if self.alive():
                return
            time.sleep(0.5)
        raise CDPError('Edge 启动失败（调试端口 %d 无响应）' % self.port)

    def new_tab(self, url='about:blank'):
        info = self._http('/json/new?' + url, method='PUT')
        return Tab(info)


class Tab(object):
    def __init__(self, info):
        self.info = info
        self.ws = websocket.create_connection(info['webSocketDebuggerUrl'], suppress_origin=True, timeout=60)
        self._id = 0
        self._pending_bodies = []
        self._req_meta = {}
        self.captured = []  # list of (url, parsed_json)
        self.capture_enabled = False

    # ---------- protocol plumbing ----------
    def send(self, method, params=None, timeout=60):
        self._id += 1
        mid = self._id
        self.ws.send(json.dumps({'id': mid, 'method': method, 'params': params or {}}))
        deadline = time.time() + timeout
        while True:
            if time.time() > deadline:
                raise CDPError('%s 超时' % method)
            msg = json.loads(self.ws.recv())
            if msg.get('id') == mid:
                if 'error' in msg:
                    raise CDPError('%s: %s' % (method, msg['error'].get('message')))
                result = msg.get('result', {})
                self._fetch_bodies()
                return result
            self._on_event(msg)

    def _on_event(self, msg):
        m = msg.get('method')
        p = msg.get('params', {})
        if m == 'Network.responseReceived':
            resp = p.get('response', {})
            mime = resp.get('mimeType', '')
            if p.get('type') in ('XHR', 'Fetch') or 'json' in mime:
                self._req_meta[p['requestId']] = resp.get('url', '')
        elif m == 'Network.loadingFinished':
            if p.get('requestId') in self._req_meta:
                self._pending_bodies.append(p['requestId'])

    def _fetch_bodies(self):
        while self._pending_bodies:
            rid = self._pending_bodies.pop(0)
            url = self._req_meta.pop(rid, '')
            try:
                self._id += 1
                mid = self._id
                self.ws.send(json.dumps({'id': mid, 'method': 'Network.getResponseBody', 'params': {'requestId': rid}}))
                while True:
                    msg = json.loads(self.ws.recv())
                    if msg.get('id') == mid:
                        break
                    self._on_event(msg)
                res = msg.get('result')
                if not res:
                    continue
                body = res.get('body', '')
                if res.get('base64Encoded'):
                    body = base64.b64decode(body).decode('utf-8', 'ignore')
                body = body.strip()
                if body[:1] in '{[':
                    try:
                        self.captured.append((url, json.loads(body)))
                    except ValueError:
                        pass
            except Exception:
                pass

    def pump(self, seconds):
        """Process network events for a while (lets JSON bodies arrive)."""
        end = time.time() + seconds
        while time.time() < end:
            self.ws.settimeout(max(0.1, min(0.5, end - time.time())))
            try:
                msg = json.loads(self.ws.recv())
                self._on_event(msg)
            except websocket.WebSocketTimeoutException:
                pass
            self.ws.settimeout(60)
            self._fetch_bodies()

    def enable_network(self):
        self.send('Network.enable', {'maxResourceBufferSize': 20000000, 'maxTotalBufferSize': 100000000})
        self.capture_enabled = True

    # ---------- page helpers ----------
    def eval(self, expr, helpers=True):
        code = '(() => {' + (JS_HELPERS if helpers else '') + 'return (' + expr + ');})()'
        r = self.send('Runtime.evaluate', {'expression': code, 'returnByValue': True, 'awaitPromise': True})
        if 'exceptionDetails' in r:
            raise CDPError('JS error: %s' % r['exceptionDetails'].get('text'))
        return r.get('result', {}).get('value')

    def activate(self):
        """Background tabs don't render or take input reliably."""
        try:
            self.send('Page.bringToFront')
        except CDPError:
            pass

    def url(self):
        try:
            return self.eval('location.href', helpers=False) or ''
        except Exception:
            return ''

    def goto(self, url, settle=4):
        self.eval('location.href = %s' % json.dumps(url), helpers=False)
        time.sleep(0.5)
        self.wait_ready()
        self.pump(settle)

    def wait_ready(self, timeout=30):
        end = time.time() + timeout
        while time.time() < end:
            try:
                if self.eval('document.readyState', helpers=False) == 'complete':
                    return True
            except Exception:
                pass
            self.pump(0.5)
        return False

    def exists(self, sel):
        return bool(self.eval('!!__q(%s)' % json.dumps(sel)))

    def wait_for(self, sel, timeout=15):
        end = time.time() + timeout
        while time.time() < end:
            if self.exists(sel):
                return True
            self.pump(0.5)
        return False

    def rect(self, sel):
        return self.eval("""(() => { let e = __q(%s); if (!e) return null;
            e.scrollIntoView({block: 'center'}); let r = e.getBoundingClientRect();
            if ((r.width === 0 || r.height === 0) && e.parentElement) r = e.parentElement.getBoundingClientRect();
            return {x: r.left + r.width / 2, y: r.top + r.height / 2}; })()""" % json.dumps(sel))

    def click(self, sel):
        """Real (trusted) mouse click at the element's center."""
        r = self.rect(sel)
        if not r:
            return False
        time.sleep(0.3)
        r = self.rect(sel) or r
        for t in ('mouseMoved', 'mousePressed', 'mouseReleased'):
            self.send('Input.dispatchMouseEvent', {'type': t, 'x': r['x'], 'y': r['y'],
                                                   'button': 'left', 'clickCount': 1})
            time.sleep(0.05)
        return True

    def type(self, sel, text):
        if not self.click(sel):
            return False
        self.eval('(() => { const e = __q(%s); if (e && e.select) e.select(); })()' % json.dumps(sel))
        self.send('Input.insertText', {'text': text})
        return True

    def is_checked(self, sel):
        return bool(self.eval('(() => { const e = __q(%s); return !!(e && e.checked); })()' % json.dumps(sel)))

    def text(self):
        return self.eval('document.body ? document.body.innerText : ""', helpers=False) or ''

    def html(self):
        return self.eval('document.documentElement.outerHTML', helpers=False) or ''

    def screenshot(self, path):
        try:
            r = self.send('Page.captureScreenshot', {'format': 'png'})
            with open(path, 'wb') as f:
                f.write(base64.b64decode(r['data']))
        except Exception:
            pass

    def scroll_bottom(self):
        self.eval("""(() => { window.scrollTo(0, document.body.scrollHeight);
            document.querySelectorAll('*').forEach(e => { if (e.scrollHeight > e.clientHeight + 50 &&
              /(auto|scroll)/.test(getComputedStyle(e).overflowY)) e.scrollTop = e.scrollHeight; }); })()""",
                  helpers=False)

    def close(self):
        try:
            self.ws.close()
        except Exception:
            pass
