#coding:utf-8

from openpyxl import load_workbook
from openpyxl.styles import Border, Side, PatternFill, Font, GradientFill, Alignment
import sys



reload(sys)
sys.setdefaultencoding('utf-8')


def reloadExcel(fileExcel):
    #读到文件
    wb = load_workbook(filename = fileExcel, data_only=True)

    #根据名称获取工作表
    #sheet_ranges = wb['1']

    #获取单元格值
    #cell_value = sheet_ranges['D18'].value

    dataDictionary = {};
    #循环所有工作表
    # for sheet in wb:
    #     print(sheet.title)

    #获取单元格区域
    # cell_range = sheet_ranges['B4':'U4']
    # for row in cell_range:
    #     for cell in row:
    #         print(cell.value)

    for sheet in wb:
        for row in sheet_ranges['A4':'U55']:
            dataList = [];
            for cell in row:
             dataList.append(cell);
            pass

    copyNewExcel();

def copyNewExcel():

    #----------------------------华丽的分隔线---------------------------------------
    #读到模板文件
    wb_template = load_workbook(filename = '个人.xlsx', data_only=True)

    #获取工作表
    template_sheet_range = wb_template['Sheet1']

    #设置单元格值
    template_sheet_range['D8'] = '这是一个测试'

    #设置工作表名称
    template_sheet_range.title = 'test'
    #复制新的工作表
    copy_sheet_range = wb_template.copy_worksheet(template_sheet_range)
    copy_sheet_range.title = 'copytest'

    #添加新的工作表
    #ws1 = wb_template.create_sheet('Mysheet')

    setSheeBoder(template_sheet_range, copy_sheet_range);

    saveExcel(wb_template);





#合并单元格样式
def style_range(ws, cell_range, border=Border(), fill=None, font=None, alignment=None):
    """
    Apply styles to a range of cells as if they were a single cell.

    :param ws:  Excel worksheet instance
    :param range: An excel range to style (e.g. A1:F20)
    :param border: An openpyxl Border
    :param fill: An openpyxl PatternFill or GradientFill
    :param font: An openpyxl Font object
    """

    top = Border(top=border.top)
    left = Border(left=border.left)
    right = Border(right=border.right)
    bottom = Border(bottom=border.bottom)

    first_cell = ws[cell_range.split(":")[0]]
    if alignment:
        ws.merge_cells(cell_range)
        first_cell.alignment = alignment

    rows = ws[cell_range]
    if font:
        first_cell.font = font

    for cell in rows[0]:
        cell.border = cell.border + top
    for cell in rows[-1]:
        cell.border = cell.border + bottom

    for row in rows:
        l = row[0]
        r = row[-1]
        l.border = l.border + left
        r.border = r.border + right
        if fill:
            for c in row:
                c.fill = fill

def setSheeBoder(template_sheet_range, copy_sheet_range):

    thin = Side(border_style="thin", color="000000")
    single = Side(border_style="thin", color="000000")

    border = Border(top=single, left=thin, right=thin, bottom=single)

    #------------------------------------华丽的分隔线------------------------------------------------                

    style_range(template_sheet_range, 'A1:K1', border=border, fill=None, font=None, alignment=None)
    style_range(template_sheet_range, 'A2:D3', border=border, fill=None, font=None, alignment=None)
    style_range(template_sheet_range, 'E2:G3', border=border, fill=None, font=None, alignment=None)
    style_range(template_sheet_range, 'H2:J3', border=border, fill=None, font=None, alignment=None)
    style_range(template_sheet_range, 'K2:K3', border=border, fill=None, font=None, alignment=None)
    style_range(template_sheet_range, 'A4:J5', border=border, fill=None, font=None, alignment=None)
    style_range(template_sheet_range, 'K4:K37', border=border, fill=None, font=None, alignment=None)

    style_range(copy_sheet_range, 'A1:K1', border=border, fill=None, font=None, alignment=None)
    style_range(copy_sheet_range, 'A2:D3', border=border, fill=None, font=None, alignment=None)
    style_range(copy_sheet_range, 'E2:G3', border=border, fill=None, font=None, alignment=None)
    style_range(copy_sheet_range, 'H2:J3', border=border, fill=None, font=None, alignment=None)
    style_range(copy_sheet_range, 'K2:K3', border=border, fill=None, font=None, alignment=None)
    style_range(copy_sheet_range, 'A4:J5', border=border, fill=None, font=None, alignment=None)
    style_range(copy_sheet_range, 'K4:K37', border=border, fill=None, font=None, alignment=None)


def saveExcel(wb_template):
    #保存工作簿 template:是否做为模板
    wb_template.template = False
    wb_template.save('document.xlsx')

def main():
    print('开始处理EXCEL----------');
    reloadExcel('TLW-3月.xlsx');

main();
