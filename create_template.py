"""生成空的投递记录工作簿。不含任何个人投递数据。"""
import os

import openpyxl
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.datavalidation import DataValidation

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, '投递记录模板.xlsx')
LAST = 301

HEADERS = ['投递日期', '平台', '公司名称', '职位名称', '城市', '薪资', '使用简历', '状态',
           '下次跟进日期', '联系人 / HR', '职位链接', '备注', '同公司同职位次数', '同公司投递次数']
WIDTHS = {'A': 12, 'B': 12, 'C': 28, 'D': 28, 'E': 10, 'F': 14, 'G': 16, 'H': 12,
          'I': 14, 'J': 16, 'K': 36, 'L': 40, 'M': 18, 'N': 16}
HEADER_FILL = PatternFill('solid', fgColor='4472C4')
RED_FILL = PatternFill('solid', fgColor='FFC7CE')
YELLOW_FILL = PatternFill('solid', fgColor='FFF2CC')
RED_FONT = Font(bold=True, color='C00000')


def build(path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = '投递记录'
    for col, name in enumerate(HEADERS, 1):
        cell = ws.cell(1, col, name)
        cell.font = Font(bold=True, color='FFFFFF')
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(vertical='center', wrap_text=True)
    for col, width in WIDTHS.items():
        ws.column_dimensions[col].width = width
    for row in range(2, LAST + 1):
        ws.cell(row, 13, '=IF(OR(C{0}="",D{0}=""),"",COUNTIFS($C:$C,C{0},$D:$D,D{0}))'.format(row))
        ws.cell(row, 14, '=IF(C{0}="","",COUNTIF($C:$C,C{0}))'.format(row))
    ws.freeze_panes = 'A2'
    ws.auto_filter.ref = 'A1:N%d' % LAST
    ws.row_dimensions[1].height = 22

    lists = {
        'B': '"Boss直聘,猎聘,智联招聘,前程无忧,内推,其他"',
        'G': '"默认简历,定制版"',
        'H': '"已打招呼,已发简历,HR已读,已查看,面试邀请,一面,二面,终面,offer,不合适,无回复,已放弃"',
    }
    for col, formula in lists.items():
        dv = DataValidation(type='list', formula1=formula, allow_blank=True)
        dv.add('%s2:%s%d' % (col, col, LAST))
        ws.add_data_validation(dv)

    data = 'A2:N%d' % LAST
    ws.conditional_formatting.add(data, FormulaRule(formula=['AND($M2<>"",$M2>1)'], fill=RED_FILL))
    ws.conditional_formatting.add(data, FormulaRule(
        formula=['AND($N2<>"",$N2>1,OR($M2="",$M2=1))'], fill=YELLOW_FILL))
    ws.conditional_formatting.add(data, FormulaRule(
        formula=['AND($I2<>"",$I2<=TODAY(),$H2<>"不合适",$H2<>"已放弃",$H2<>"offer")'], font=RED_FONT))

    stat = wb.create_sheet('统计')
    stat['A1'] = '按平台'
    stat['A1'].font = Font(bold=True)
    for i, name in enumerate(['Boss直聘', '猎聘', '智联招聘', '前程无忧', '内推', '其他'], 2):
        stat.cell(i, 1, name)
        stat.cell(i, 2, '=COUNTIF(投递记录!$B:$B,A%d)' % i)
    stat['A9'] = '按状态'
    stat['A9'].font = Font(bold=True)
    statuses = ['已打招呼', '已发简历', 'HR已读', '已查看', '面试邀请', '一面', '二面', '终面',
                'offer', '不合适', '无回复', '已放弃']
    for i, name in enumerate(statuses, 10):
        stat.cell(i, 1, name)
        stat.cell(i, 2, '=COUNTIF(投递记录!$H:$H,A%d)' % i)
    stat['A23'] = '总投递数'
    stat['B23'] = '=COUNTA(投递记录!$C$2:$C$301)'
    stat['A24'] = '重复投递条数'
    stat['B24'] = '=COUNTIF(投递记录!$M$2:$M$301,">1")'
    stat['A25'] = '今天需要跟进'
    stat['B25'] = ('=COUNTIFS(投递记录!$I$2:$I$301,"<="&TODAY(),投递记录!$H$2:$H$301,"<>不合适",'
                   '投递记录!$H$2:$H$301,"<>已放弃",投递记录!$H$2:$H$301,"<>offer")')
    stat.column_dimensions['A'].width = 18
    stat.column_dimensions['B'].width = 14

    info = wb.create_sheet('使用说明')
    lines = [
        '使用说明',
        '1. 这是空模板。请复制到别的位置再使用，不要把投递数据写回程序目录里的这一份。',
        '2. 平台、使用简历、状态都有下拉选项。使用简历里只有「默认简历、定制版」，可以在 Excel 的数据验证里改成你自己的简历名称。',
        '3. 程序同步时按「平台 + 公司 + 职位」去重。公司名尽量写全称，或固定用同一种简称。',
        '4. 整行红色：同一公司、同一职位已经记录过不止一次。整行黄色：同一公司投过不同职位。',
        '5. 字体红色加粗：「下次跟进日期」已到或已过，且状态不是不合适、已放弃、offer。',
        '6. 「不合适」状态格为红色，「已查看」状态格为黄色，这是同步程序写进去的。',
        '7. 「统计」表自动汇总。同步程序还会追加一张「同步日志」。',
    ]
    for i, text in enumerate(lines, 1):
        info.cell(i, 1, text)
        if i == 1:
            info.cell(i, 1).font = Font(bold=True, size=14)
    info.column_dimensions['A'].width = 110
    wb.save(path)


if __name__ == '__main__':
    build(OUT)
    print(OUT)
