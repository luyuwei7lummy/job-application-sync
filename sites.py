"""Per-site login + scraping logic for Boss直聘 / 猎聘 / 智联招聘."""
import json
import os
import re
import time

import excel_sync
from extract import walk

SITES = [
    {
        'key': 'boss',
        'platform': 'Boss直聘',
        'list_url': 'https://www.zhipin.com/web/geek/recommend',
        'login_url': 'https://www.zhipin.com/web/user/?ka=header-login',
        'login_url_re': r'zhipin\.com/web/user',
        'verify_url_re': r'verify\.html|/web/passport/zp/verify',
        'logged_out_sel': 'a[ka="header-login"]',
        'phone_sel': 'input[type="tel"][placeholder="手机号"]',
        'agree_sel': 'input.agree-policy',
        'send_sel': '.btn-sms',
        'code_sel': 'input[placeholder="短信验证码"]',
        'submit_sel': 'button.sure-btn',
        # 列表页 URL 自带页码，逐页打开即可；感兴趣(tag=2) 侧栏每页都会加载，必须按 tag 过滤
        'views': [
            {'name': '沟通过', 'status': '已打招呼', 'use_text': False, 'title_filter': True,
             'page_url': 'https://www.zhipin.com/web/geek/recommend?tab=1&page={page}&tag=5',
             'api_include': r'/interaction/geekGetJob\?.*tag=5'},
            {'name': '已投递', 'status': '已发简历', 'use_text': False,
             'page_url': 'https://www.zhipin.com/web/geek/recommend?tab=2&page={page}&tag=3',
             'api_include': r'/resume/geekDeliverList'},
        ],
        'link_fmt': 'https://www.zhipin.com/job_detail/{}.html',
    },
    {
        'key': 'liepin',
        'platform': '猎聘',
        'list_url': 'https://c.liepin.com/job/record/apply',
        'login_url': None,  # 未登录时自动跳到首页，登录框就在首页
        'login_url_re': r'^https?://www\.liepin\.com/.*loginBackUrl|passport\.liepin\.com',
        'verify_url_re': r'^about:blank$|captcha|verify',
        'logged_out_sel': None,
        'phone_sel': 'input[placeholder="手机号"]',
        'agree_sel': '.c-login-agreement-box input[type="checkbox"]',
        'send_sel': '.get-smscode-btn',
        'code_sel': 'input[placeholder="短信验证码"]',
        'submit_sel': 'button.login-submit-btn',
        'views': [
            {'name': '全部状态', 'tab_text': '^全部状态$', 'status': '已发简历', 'use_text': True,
             'api_include': r'capply\.pc\.apply-list'},
        ],
        'link_fmt': 'https://www.liepin.com/job/{}.shtml',
    },
    {
        'key': 'zhaopin',
        'platform': '智联招聘',
        'list_url': 'https://i.zhaopin.com/schedule',
        'login_url': None,
        'login_url_re': r'passport\.zhaopin\.com',
        'verify_url_re': r'captcha|verify',
        'logged_out_sel': None,
        'phone_sel': 'input.zppp-input[type="tel"]',
        'agree_sel': 'input.zppp-accept__checkbox',
        'send_sel': 'button.zppp-sms__send',
        'code_sel': 'input[placeholder="短信验证码"]',
        'submit_sel': 'button.zppp-submit',
        'views': [
            {'name': '投递成功', 'tab_text': '^(投递成功|成功投递)', 'status': '已发简历', 'use_text': False,
             'api_include': r'/schedule/feedback\?.*status=send'},
            {'name': '不合适', 'tab_text': '^不合适', 'status': '不合适', 'use_text': False,
             'api_include': r'/schedule/feedback\?.*status=unsuitable'},
        ],
        'link_fmt': 'https://jobs.zhaopin.com/{}.htm',
    },
]


def map_status(text):
    t = text or ''
    if re.search(r'不合适|未通过|未录取|拒绝|淘汰|不匹配', t):
        return '不合适'
    if re.search(r'offer|录用', t, re.I):
        return 'offer'
    if re.search(r'面试', t):
        return '面试邀请'
    if re.search(r'未查看|待查看', t):
        return '已发简历'
    if re.search(r'查看|已读|感兴趣|有意向|沟通中|看过|筛选', t):
        return '已查看'
    if re.search(r'投递|成功|待处理|已发|待查看|未查看', t):
        return '已发简历'
    return ''


NEXT_JS = r"""(() => {
  const cands = Array.from(document.querySelectorAll('button,a,li,span,div,i')).filter(e => {
    if (!__vis(e)) return false;
    const t = (e.innerText || '').trim(); const c = (e.className || '').toString();
    const a = e.getAttribute('aria-label') || e.getAttribute('title') || '';
    return /^(下一页|加载更多|查看更多)$/.test(t) || /下一页|next page/i.test(a) ||
      (/(^|[\s_-])(next|btn-next|ant-pagination-next|ui-pager-next)([\s_-]|$)/.test(c) && t.length < 6);
  });
  const ok = cands.find(e => { const c = (e.className || '').toString() + ' ' + ((e.parentElement || {}).className || '');
    return !/disabled/.test(c) && !e.disabled && e.getAttribute('aria-disabled') !== 'true'; });
  if (!ok) return null;
  ok.scrollIntoView({block: 'center'}); const r = ok.getBoundingClientRect();
  return {x: r.left + r.width / 2, y: r.top + r.height / 2};
})()"""


class Site(object):
    def __init__(self, cfg, tab, ui, phone, log, dump_dir=None, max_pages=30, include_kw=(), exclude_kw=()):
        self.c = cfg
        self.include_kw = include_kw
        self.exclude_kw = exclude_kw
        self.tab = tab
        self.ui = ui
        self.phone = phone
        self.log = log
        self.dump_dir = dump_dir
        self.max_pages = max_pages

    # ---------- login ----------
    def _logged_out(self):
        url = self.tab.url()
        if re.search(self.c['login_url_re'], url):
            return True
        if self.c['logged_out_sel'] and self.tab.exists(self.c['logged_out_sel']):
            return True
        return False

    def _handle_verify(self):
        for _ in range(3):
            url = self.tab.url()
            if not re.search(self.c['verify_url_re'], url):
                return
            if url == 'about:blank':
                self.tab.goto(self.c['list_url'], settle=4)
                continue
            self.ui.info('%s 出现了安全验证页面。\n\n请在浏览器里点按钮/拖滑块完成验证，完成后点「确定」继续。'
                         % self.c['platform'])
            self.tab.pump(3)

    def ensure_login(self):
        p = self.c['platform']
        self.tab.activate()
        self.log('[%s] 打开 %s' % (p, self.c['list_url']))
        self.tab.goto(self.c['list_url'], settle=4)
        self._handle_verify()
        if not self._logged_out():
            self.log('[%s] 已是登录状态' % p)
            return True
        for attempt in range(2):
            self.log('[%s] 需要登录，开始手机号验证码登录' % p)
            self._login_once()
            self.tab.goto(self.c['list_url'], settle=4)
            self._handle_verify()
            if not self._logged_out():
                self.log('[%s] 登录成功' % p)
                return True
        self.ui.info('%s 仍未登录。请在浏览器里手动登录，完成后点「确定」。' % p)
        self.tab.goto(self.c['list_url'], settle=4)
        return not self._logged_out()

    def _login_once(self):
        t, c = self.tab, self.c
        if c['login_url'] and not re.search(c['login_url_re'], t.url()):
            t.goto(c['login_url'], settle=3)
            self._handle_verify()
        if not t.wait_for(c['phone_sel'], 15):
            self.ui.info('%s 没找到手机号输入框，请在浏览器里手动登录，完成后点「确定」。' % c['platform'])
            return
        t.type(c['phone_sel'], self.phone)
        time.sleep(0.5)
        self._agree()
        t.click(c['send_sel'])
        t.pump(1.5)
        code = self.ui.ask_code(
            '%s 已点击「发送验证码」到 %s。\n\n如果浏览器弹出滑块/点选验证，请先在浏览器里完成它。\n'
            '收到短信后在下面输入验证码（点取消=我在浏览器里自己登录）：' % (c['platform'], self.phone))
        if not code:
            self.ui.info('请在浏览器里完成 %s 登录，完成后点「确定」。' % c['platform'])
            return
        t.type(c['code_sel'], code.strip())
        time.sleep(0.5)
        self._agree()
        t.click(c['submit_sel'])
        for _ in range(20):
            t.pump(1)
            if not re.search(c['login_url_re'], t.url()) or not t.exists(c['code_sel']):
                return
        self.ui.info('%s 自动提交后似乎还停在登录页，请在浏览器里检查/完成登录，然后点「确定」。' % c['platform'])

    def _agree(self):
        t, sel = self.tab, self.c['agree_sel']
        if not t.exists(sel) or t.is_checked(sel):
            return
        t.click(sel)
        time.sleep(0.4)
        if not t.is_checked(sel):
            t.click('text=^已阅读并同意')
            time.sleep(0.4)

    # ---------- scraping ----------
    def _collect(self, view):
        out = {}
        inc = re.compile(view['api_include'])
        for url, data in self.tab.captured:
            if not inc.search(url):
                continue
            for rec in walk(data):
                key = (excel_sync.norm_company(rec['company']), excel_sync.norm(rec['title']))
                out.setdefault(key, rec)
        return out

    def _next_page(self):
        r = self.tab.eval(NEXT_JS)
        if not r:
            return False
        for typ in ('mouseMoved', 'mousePressed', 'mouseReleased'):
            self.tab.send('Input.dispatchMouseEvent', {'type': typ, 'x': r['x'], 'y': r['y'],
                                                       'button': 'left', 'clickCount': 1})
        return True

    def scrape_view(self, view):
        p = self.c['platform']
        t = self.tab
        t.activate()
        t.captured = []
        if view.get('page_url'):
            recs = {}
            for page in range(1, self.max_pages + 1):
                before = len(recs)
                t.goto(view['page_url'].format(page=page), settle=3)
                self._handle_verify()
                recs = self._collect(view)
                if len(recs) == before:
                    break
                time.sleep(1)
        else:
            # 重新打开列表页，保证第一页的接口数据被抓到
            t.goto(self.c['list_url'], settle=3)
            self._handle_verify()
            if view.get('tab_text'):
                if t.click('text=' + view['tab_text']):
                    self.log('[%s] 切换到「%s」' % (p, view['name']))
                else:
                    self.log('[%s] 没找到「%s」标签，按当前页面抓取' % (p, view['name']))
            t.pump(3)
            recs = self._collect(view)
            stale = 0
            for _ in range(self.max_pages):
                before = len(recs)
                t.scroll_bottom()
                t.pump(1.5)
                moved = self._next_page()
                if moved:
                    t.pump(3)
                recs = self._collect(view)
                stale = stale + 1 if len(recs) == before else 0
                if stale >= 2 or (not moved and stale >= 1):
                    break
        self.log('[%s]「%s」抓到 %d 条' % (p, view['name'], len(recs)))
        if view.get('title_filter') and self.include_kw:
            kept = {k: r for k, r in recs.items() if self._title_ok(r['title'])}
            self.log('[%s]「%s」按职位关键词过滤后保留 %d 条（跳过 %d 条无关职位）'
                     % (p, view['name'], len(kept), len(recs) - len(kept)))
            recs = kept
        if not recs:
            self.log('[%s]「%s」没抓到数据，页面快照已存到 debug 目录' % (p, view['name']))
        self._dump(view)
        result = []
        for rec in recs.values():
            mapped = map_status(rec['status_text']) if view.get('use_text') else ''
            rec['status'] = mapped or view['status']
            rec['platform'] = p
            if not rec['link'] and rec.get('job_id') and self.c.get('link_fmt'):
                rec['link'] = self.c['link_fmt'].format(rec['job_id'])
            if rec.get('hr') and rec.get('hr_title'):
                rec['hr'] = '%s（%s）' % (rec['hr'], rec['hr_title'])
            result.append(rec)
        return result

    def _title_ok(self, title):
        t = (title or '').lower()
        if any(k.lower() in t for k in self.exclude_kw):
            return False
        return any(k.lower() in t for k in self.include_kw)

    def _dump(self, view):
        if not self.dump_dir:
            return
        base = os.path.join(self.dump_dir, '%s_%s' % (self.c['key'], view['name']))
        try:
            with open(base + '.json', 'w', encoding='utf-8') as f:
                json.dump([{'url': u, 'data': d} for u, d in self.tab.captured], f, ensure_ascii=False, indent=1)
            with open(base + '.txt', 'w', encoding='utf-8') as f:
                f.write(self.tab.url() + '\n\n' + self.tab.text())
            self.tab.screenshot(base + '.png')
        except Exception as e:
            self.log('调试数据保存失败: %s' % e)

    def run(self):
        if not self.ensure_login():
            self.log('[%s] 未登录，跳过' % self.c['platform'])
            return []
        merged = {}
        for view in self.c['views']:
            for rec in self.scrape_view(view):
                key = (excel_sync.norm_company(rec['company']), excel_sync.norm(rec['title']))
                old = merged.get(key)
                if not old or excel_sync.should_update(old['status'], rec['status']):
                    merged[key] = rec
        return list(merged.values())
