"""Merge scraped application records into 求职投递记录.xlsx."""
import datetime
import os
import re
import shutil

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill

SHEET = '投递记录'
LOG_SHEET = '同步日志'
COL = {'date': 1, 'platform': 2, 'company': 3, 'title': 4, 'city': 5, 'salary': 6, 'resume': 7,
       'status': 8, 'follow': 9, 'hr': 10, 'link': 11, 'note': 12}
FORMULA_LAST_ROW = 301

RANK = {'已打招呼': 1, '已发简历': 2, 'HR已读': 3, '已查看': 3, '面试邀请': 4, '一面': 5, '二面': 6,
        '终面': 7, 'offer': 8}
TERMINAL = ('不合适', '已放弃', '无回复')
# 只给状态列上色：不合适为红，已查看为黄。其它状态不涂，避免和这两种含义混在一起。
STATUS_FILL = {
    '不合适': PatternFill('solid', fgColor='FF6B6B'),
    '已查看': PatternFill('solid', fgColor='FFE566'),
    'HR已读': PatternFill('solid', fgColor='FFE566'),
}
STATUS_FILL_RGB = {'不合适': 'FF6B6B', '已查看': 'FFE566', 'HR已读': 'FFE566'}
# 以前状态一变化就涂成 FFF2CC。现在黄色只表示「已查看」，所以顺手清掉这种旧底色。
OWNED_RGB = set(STATUS_FILL_RGB.values()) | {'FFF2CC'}

_COMPANY_SUFFIX = re.compile(r'(股份有限公司|有限责任公司|有限公司|集团|股份|控股|科技|技术|分公司)')


def norm(s):
    s = (s or '').strip().lower()
    s = s.replace('（', '(').replace('）', ')').replace('【', '[').replace('】', ']')
    s = re.sub(r'\[[^\]]*\]', '', s)
    s = re.sub(r'\s+', '', s)
    return s


def norm_company(s):
    s = norm(s)
    s = re.sub(r'\([^)]*\)', '', s)
    return _COMPANY_SUFFIX.sub('', s)


def same_company(a, b):
    a, b = norm_company(a), norm_company(b)
    if not a or not b:
        return False
    return a == b or (len(a) >= 2 and len(b) >= 2 and (a in b or b in a))


def same_title(a, b):
    a, b = norm(a), norm(b)
    if a == b:
        return True
    ca = re.search(r'\(?(j\d{4,})\)?', a)
    cb = re.search(r'\(?(j\d{4,})\)?', b)
    if ca and cb:
        return ca.group(1) == cb.group(1)
    strip = lambda x: re.sub(r'\([^)]*\)', '', x)
    if strip(a) == strip(b) and strip(a) != '':
        return True
    short, long_ = sorted((a, b), key=len)
    return len(short) >= 6 and long_.startswith(short)


def _fill_rgb(cell):
    fill = cell.fill
    if not fill or fill.fill_type != 'solid' or fill.fgColor is None:
        return None
    if fill.fgColor.type != 'rgb' or not fill.fgColor.rgb:
        return None
    return str(fill.fgColor.rgb)[-6:].upper()


def paint_status(cell):
    """按状态给单元格上色。颜色已经对上时不改。返回是否改动了填充。"""
    status = cell.value or ''
    wanted = STATUS_FILL_RGB.get(status)
    current = _fill_rgb(cell)
    if wanted:
        if current == wanted:
            return False
        cell.fill = STATUS_FILL[status]
        return True
    if current in OWNED_RGB:
        cell.fill = PatternFill(fill_type=None)
        return True
    return False


def color_status_column(ws, rows):
    changed = False
    for row in rows:
        if paint_status(ws.cell(row=row, column=COL['status'])):
            changed = True
    return changed


def should_update(old, new):
    if not new or new == old:
        return False
    if not old:
        return True
    if new == '不合适':
        return old not in ('不合适', 'offer', '已放弃')
    if old in ('offer', '已放弃'):
        return False
    if old in TERMINAL:
        return new in RANK and new != '已打招呼'
    return RANK.get(new, 0) > RANK.get(old, 0)


class SaveFailed(Exception):
    pass


def merge(path, records, ask_retry, log):
    """records: list of dicts with platform/company/title/status/... ; returns summary dict."""
    wb = openpyxl.load_workbook(path)
    ws = wb[SHEET]
    rows = []
    r = 2
    while ws.cell(row=r, column=COL['company']).value or ws.cell(row=r, column=COL['title']).value:
        rows.append(r)
        r += 1
    next_row = r

    if LOG_SHEET in wb.sheetnames:
        lws = wb[LOG_SHEET]
    else:
        lws = wb.create_sheet(LOG_SHEET)
        for i, h in enumerate(['同步时间', '平台', '公司名称', '职位名称', '变化', '原状态', '新状态'], 1):
            c = lws.cell(row=1, column=i, value=h)
            c.font = Font(bold=True, color='FFFFFF')
            c.fill = PatternFill('solid', fgColor='4472C4')
        for col, w in zip('ABCDEFG', (18, 10, 30, 34, 10, 10, 10)):
            lws.column_dimensions[col].width = w
        lws.freeze_panes = 'A2'
    now = datetime.datetime.now().strftime('%Y-%m-%d %H:%M')
    today = datetime.date.today()

    summary = {'new': [], 'changed': [], 'same': 0}
    for rec in records:
        platform = rec['platform']
        hit = None
        for row in rows:
            if ws.cell(row=row, column=COL['platform']).value != platform:
                continue
            if same_title(ws.cell(row=row, column=COL['title']).value, rec['title']) and \
                    same_company(ws.cell(row=row, column=COL['company']).value, rec['company']):
                hit = row
                break
        if hit:
            old = ws.cell(row=hit, column=COL['status']).value or ''
            new = rec.get('status') or ''
            for key in ('city', 'salary', 'link', 'hr'):
                if rec.get(key) and not ws.cell(row=hit, column=COL[key]).value:
                    ws.cell(row=hit, column=COL[key], value=rec[key])
            if rec.get('date') and not ws.cell(row=hit, column=COL['date']).value:
                c = ws.cell(row=hit, column=COL['date'], value=rec['date'])
                c.number_format = 'yyyy-mm-dd'
            if should_update(old, new):
                ws.cell(row=hit, column=COL['status'], value=new)
                note = ws.cell(row=hit, column=COL['note']).value or ''
                add = '%s 同步：%s→%s' % (today.strftime('%m-%d'), old or '空', new)
                ws.cell(row=hit, column=COL['note'], value=(note + '；' + add) if note else add)
                lws.append([now, platform, rec['company'], rec['title'], '状态变化', old, new])
                summary['changed'].append((platform, rec['company'], rec['title'], old, new))
            else:
                summary['same'] += 1
            continue

        row = next_row
        next_row += 1
        rows.append(row)
        vals = {'date': rec.get('date') or None, 'platform': platform, 'company': rec['company'],
                'title': rec['title'], 'city': rec.get('city') or None, 'salary': rec.get('salary') or None,
                'status': rec.get('status') or None, 'hr': rec.get('hr') or None,
                'link': rec.get('link') or None,
                'note': ('%s 自动同步新增' % today.strftime('%m-%d')) + (
                    '；平台状态：%s' % rec['status_text'] if rec.get('status_text') and
                    rec['status_text'] != rec.get('status') else '')}
        for key, v in vals.items():
            c = ws.cell(row=row, column=COL[key], value=v)
            c.alignment = Alignment(vertical='center', wrap_text=key in ('title', 'note'))
            if key == 'date' and v:
                c.number_format = 'yyyy-mm-dd'
        if row > FORMULA_LAST_ROW:
            ws.cell(row=row, column=13,
                    value='=IF(OR(C{0}="",D{0}=""),"",COUNTIFS($C:$C,C{0},$D:$D,D{0}))'.format(row))
            ws.cell(row=row, column=14, value='=IF(C{0}="","",COUNTIF($C:$C,C{0}))'.format(row))
        lws.append([now, platform, rec['company'], rec['title'], '新增', '', rec.get('status') or ''])
        summary['new'].append((platform, rec['company'], rec['title'], rec.get('status')))

    colored = color_status_column(ws, rows)
    if summary['new'] or summary['changed'] or colored:
        save(wb, path, ask_retry, log)
    return summary


def save(wb, path, ask_retry, log):
    bdir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'backups')
    os.makedirs(bdir, exist_ok=True)
    try:
        shutil.copy2(path, os.path.join(bdir, '投递记录_' + datetime.datetime.now().strftime('%Y%m%d_%H%M%S') + '.xlsx'))
    except OSError:
        pass
    while True:
        try:
            wb.save(path)
            log('已写入 %s' % path)
            return
        except PermissionError:
            if not ask_retry('Excel 文件正被打开（%s）。\n请关闭 Excel 后点「重试」；点「取消」则另存为新文件。'
                             % os.path.basename(path)):
                alt = path.replace('.xlsx', '_同步_' + datetime.datetime.now().strftime('%H%M%S') + '.xlsx')
                wb.save(alt)
                log('原文件被占用，已另存为 %s' % alt)
                return
