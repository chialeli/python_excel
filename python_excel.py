# coding=utf-8
import xlrd
import xlwt

data = xlrd.open_workbook('成绩单.xls')

sheet_num = data.nsheets
print('所有表格数量：{}'.format(sheet_num))

sheet_name = data.sheet_names()[0]  # 获取所有sheet名称
print(sheet_name)

sheet_name = data.sheet_names()[1]  # 获取所有sheet名称
print(sheet_name)

sheet = data.sheet_by_index(0)
print('表格名称:{}\n表格列数: {}\n表格行数: {}'.format(sheet.name, sheet.ncols, sheet.nrows))

for i in range(sheet.nrows):
    for j in range(sheet.ncols):
        print(sheet.cell(i,j).value)
