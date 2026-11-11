"""求职投递同步：登录 Boss直聘 / 猎聘 / 智联招聘，抓取投递记录并合并到 Excel。"""
import argparse
import datetime
import json
import os
import re
import sys
import traceback
import tkinter as tk
from tkinter import messagebox, simpledialog

import excel_sync
from cdp import Browser
from sites import SITES, Site

HERE = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(HERE, 'config.json')
TEMPLATE_PATH = os.path.join(HERE, 'config.template.json')

# 只有空位和说明，没有任何手机号、路径或别人的关键词。
# 程序只在缺少 config.json 时把这份内容写成 config.template.json，不会替用户填，也不会拿它直接跑。
CONFIG_TEMPLATE = {
    '_说明': '把本文件复制为 config.json 后再填写。程序不带任何默认手机号或路径，空着就拒绝运行。以 _ 开头的键只给人看，程序不读取。不要使用别人的 config.json。',
    '_phone': '填你自己的 11 位手机号，三个网站都用这个号登录。',
    'phone': '',
    '_excel': '投递记录工作簿的完整路径。文件必须已经存在，且有一个叫「投递记录」的工作表。运行前先关掉 Excel。',
    'excel': '',
    '_edge': '本机 msedge.exe 的完整路径。可在开始菜单里对 Microsoft Edge 点右键，打开文件所在位置查看。',
    'edge': '',
    '_port': 'Edge 调试端口，填一个 1 到 65535 的整数。本机这个端口没被占用即可，常见写法是 9333。',
    'port': None,
    '_max_pages': '每个列表最多翻多少页，填正整数。',
    'max_pages': None,
    '_debug_dump': '填 true 或 false。true 时把当次接口数据和截图存到 debug 目录，方便排查抓不到数据的问题。',
    'debug_dump': None,
    '_boss_contact_include': '只作用于 Boss「沟通过」。职位名包含其中任意一个词才导入。留空数组表示不过滤、全部导入。',
    'boss_contact_include': [],
    '_boss_contact_exclude': '只作用于 Boss「沟通过」。职位名包含其中任意一个词就丢弃。留空数组表示没有排除词。',
    'boss_contact_exclude': [],
}


class ConfigError(Exception):
    pass


def write_template():
    with open(TEMPLATE_PATH, 'w', encoding='utf-8') as f:
        json.dump(CONFIG_TEMPLATE, f, ensure_ascii=False, indent=2)
        f.write('\n')


def load_config():
    """读取 config.json。没有这个文件就只生成空模板并停止，绝不套用默认值。"""
    if not os.path.exists(CONFIG_PATH):
        write_template()
        raise ConfigError(
            '没有找到 config.json，程序已停止。\n\n'
            '已在程序目录生成 config.template.json。\n'
            '请把它复制一份，改名为 config.json，用记事本按里面的说明填上你自己的手机号、Excel 路径和 Edge 路径，保存后再运行。\n\n'
            '程序不会自动填任何默认值。不要把别人的 config.json 拷过来用。')
    try:
        with open(CONFIG_PATH, encoding='utf-8') as f:
            raw = json.load(f)
    except ValueError as e:
        raise ConfigError('config.json 不是合法的 JSON，请检查逗号和引号。\n%s' % e)
    if not isinstance(raw, dict):
        raise ConfigError('config.json 的最外层必须是 { }。')
    cfg = dict((k, v) for k, v in raw.items() if not str(k).startswith('_'))
    errors = _validate(cfg)
    if errors:
        raise ConfigError('config.json 还不能用，程序已停止，没有改这个文件：\n\n- ' + '\n- '.join(errors) +
                          '\n\n按 config.template.json 里的说明改完并保存后，再重新运行。')
    return cfg


def _validate(cfg):
    errors = []
    required = ('phone', 'excel', 'edge', 'port', 'max_pages', 'debug_dump',
                'boss_contact_include', 'boss_contact_exclude')
    for key in required:
        if key not in cfg:
            errors.append('缺少 %s' % key)

    phone = cfg.get('phone')
    if 'phone' in cfg and (not isinstance(phone, str) or not re_phone(phone)):
        errors.append('phone 必须是你自己的 11 位手机号，现在是空的或格式不对')

    excel = cfg.get('excel')
    if 'excel' in cfg:
        if not isinstance(excel, str) or not excel.strip():
            errors.append('excel 是空的，请填投递记录工作簿的完整路径')
        elif not os.path.isfile(excel):
            errors.append('excel 指向的文件不存在：%s。请先把程序目录里的「投递记录模板.xlsx」复制到你自己的位置，再把这个路径填上' % excel)
        else:
            try:
                import openpyxl
                wb = openpyxl.load_workbook(excel, read_only=True)
                try:
                    if '投递记录' not in wb.sheetnames:
                        errors.append('excel 里没有「投递记录」这个工作表。请使用「投递记录模板.xlsx」，或在工作簿里建一个同名工作表')
                finally:
                    wb.close()
            except Exception as e:
                errors.append('excel 打不开：%s' % e)

    edge = cfg.get('edge')
    if 'edge' in cfg:
        if not isinstance(edge, str) or not edge.strip():
            errors.append('edge 是空的，请填 msedge.exe 的完整路径')
        elif not os.path.isfile(edge):
            errors.append('edge 指向的文件不存在：%s' % edge)

    port = cfg.get('port')
    if 'port' in cfg and not (isinstance(port, int) and not isinstance(port, bool) and 1 <= port <= 65535):
        errors.append('port 必须是 1 到 65535 的整数，不能留空')

    pages = cfg.get('max_pages')
    if 'max_pages' in cfg and not (isinstance(pages, int) and not isinstance(pages, bool) and pages >= 1):
        errors.append('max_pages 必须是大于 0 的整数，不能留空')

    if 'debug_dump' in cfg and not isinstance(cfg.get('debug_dump'), bool):
        errors.append('debug_dump 必须是 true 或 false，不能留空')

    for key in ('boss_contact_include', 'boss_contact_exclude'):
        value = cfg.get(key)
        if key in cfg and not (isinstance(value, list) and all(isinstance(x, str) and x.strip() for x in value)):
            errors.append('%s 必须是字符串数组，不需要过滤时写成 []' % key)
    return errors


def re_phone(phone):
    return bool(re.fullmatch(r'1\d{10}', (phone or '').strip()))


class UI(object):
    def __init__(self):
        self.root = tk.Tk()
        self.root.withdraw()
        self.root.title('求职投递同步')
        self.root.attributes('-topmost', True)

    def choose(self):
        """Start window. Returns (site_keys, dry_run) or None."""
        r = self.root
        r.deiconify()
        tk.Label(r, text='选择要同步的网站：', font=('Microsoft YaHei', 11)).pack(anchor='w', padx=16, pady=(14, 4))
        vars_ = {}
        for s in SITES:
            v = tk.BooleanVar(value=True)
            tk.Checkbutton(r, text=s['platform'], variable=v, font=('Microsoft YaHei', 10)).pack(anchor='w', padx=28)
            vars_[s['key']] = v
        dry = tk.BooleanVar(value=False)
        tk.Checkbutton(r, text='预览模式（只抓取，不写 Excel）', variable=dry,
                       font=('Microsoft YaHei', 10)).pack(anchor='w', padx=16, pady=(8, 0))
        result = {}

        def go():
            result['keys'] = [k for k, v in vars_.items() if v.get()]
            result['dry'] = dry.get()
            r.quit()

        tk.Button(r, text='开始同步', width=16, command=go, font=('Microsoft YaHei', 10)).pack(pady=14)
        r.protocol('WM_DELETE_WINDOW', r.quit)
        r.geometry('+200+200')
        r.mainloop()
        for w in r.winfo_children():
            w.destroy()
        r.withdraw()
        if 'keys' not in result:
            return None
        return result['keys'], result['dry']

    def ask_code(self, msg):
        return simpledialog.askstring('输入短信验证码', msg, parent=self.root)

    def info(self, msg):
        messagebox.showinfo('求职投递同步', msg, parent=self.root)

    def ask_retry(self, msg):
        return messagebox.askretrycancel('求职投递同步', msg, parent=self.root)


def main():
    ap = argparse.ArgumentParser(description='同步 Boss直聘/猎聘/智联 投递记录到 Excel')
    ap.add_argument('--sites', help='逗号分隔：boss,liepin,zhaopin；给了就跳过开始窗口')
    ap.add_argument('--preview', action='store_true', help='只抓取不写 Excel')
    args = ap.parse_args()

    ui = UI()
    try:
        cfg = load_config()
    except ConfigError as e:
        print(e)
        ui.info(str(e))
        return
    stamp = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
    os.makedirs(os.path.join(HERE, 'logs'), exist_ok=True)
    logf = open(os.path.join(HERE, 'logs', 'sync_%s.log' % stamp), 'w', encoding='utf-8')

    def log(msg):
        line = '%s %s' % (datetime.datetime.now().strftime('%H:%M:%S'), msg)
        print(line)
        logf.write(line + '\n')
        logf.flush()

    if args.sites:
        ui.root.withdraw()
        keys, dry = [k.strip() for k in args.sites.split(',')], args.preview
    else:
        choice = ui.choose()
        if not choice:
            return
        keys, dry = choice
    dump_dir = None
    if cfg.get('debug_dump'):
        dump_dir = os.path.join(HERE, 'debug', stamp)
        os.makedirs(dump_dir, exist_ok=True)

    browser = Browser(cfg['edge'], os.path.join(HERE, 'edge_profile'), cfg['port'])
    browser.start()
    log('Edge 已就绪（独立配置目录，登录状态会保留到下次）')

    records, errors = [], []
    for sc in SITES:
        if sc['key'] not in keys:
            continue
        tab = None
        try:
            tab = browser.new_tab('about:blank')
            tab.enable_network()
            site = Site(sc, tab, ui, cfg['phone'], log, dump_dir, cfg['max_pages'],
                        cfg.get('boss_contact_include', ()), cfg.get('boss_contact_exclude', ()))
            got = site.run()
            log('[%s] 合计 %d 条（已去重）' % (sc['platform'], len(got)))
            records.extend(got)
        except Exception as e:
            errors.append('%s: %s' % (sc['platform'], e))
            log('[%s] 出错: %s\n%s' % (sc['platform'], e, traceback.format_exc()))
        finally:
            if tab:
                tab.close()

    if dry:
        for r in records:
            log('  预览 | %s | %s | %s | %s | %s' % (r['platform'], r['company'], r['title'],
                                                   r['status'], r.get('status_text', '')))
        msg = '预览完成：共抓到 %d 条记录，未写入 Excel。\n详细见窗口输出 / logs 目录。' % len(records)
    else:
        summary = excel_sync.merge(cfg['excel'], records, ui.ask_retry, log)
        lines = ['同步完成：新增 %d 条，状态变化 %d 条，无变化 %d 条。' %
                 (len(summary['new']), len(summary['changed']), summary['same'])]
        if summary['changed']:
            lines.append('\n状态变化：')
            lines += ['  [%s] %s / %s：%s → %s' % c for c in summary['changed'][:15]]
        if summary['new']:
            lines.append('\n新增：')
            lines += ['  [%s] %s / %s（%s）' % n for n in summary['new'][:15]]
            if len(summary['new']) > 15:
                lines.append('  …… 其余见 Excel「同步日志」表')
        msg = '\n'.join(lines)
        log(msg)
    if errors:
        msg += '\n\n出错：\n' + '\n'.join(errors)
        log('出错：' + '；'.join(errors))
    if not args.sites:
        ui.info(msg)
    logf.close()


if __name__ == '__main__':
    try:
        main()
    except Exception:
        traceback.print_exc()
        input('出错了，按回车退出')
        sys.exit(1)
