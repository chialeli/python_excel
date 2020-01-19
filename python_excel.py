# coding=utf-8
import xlrd
import xlwt

read_sheets():
    data = xlrd.open_workbook('成绩单.xls')
    print('所有表格数量：{}'.format(data.nsheets))
    for sheet_num in range(data.nsheets):
        sheet_name = data.sheet_names()[sheet_num]
        print('\n')
        print(sheet_name)
        print("#########")

        sheet = data.sheet_by_index(sheet_num)
        print('表格名称:{}\n表格列数: {}\n表格行数: {}'.format(sheet.name, sheet.ncols, sheet.nrows))
        for col_num in range(sheel.ncols):
            col_name = sheel.cell(0, col_num)self.
        for row_num in range(1, sheet.nrows):
            for col_num in range(sheet.ncols):
                value = sheet.cell(row_num, col_num).value
                #if (len(sheet.cell(j, k)) != 0):
                print(sheet.cell(row_num, col_num).value)
                #else: 
                #    print("null")
write_sheets():


if __name == "__main__"
    read_sheets()
    write_sheets()

