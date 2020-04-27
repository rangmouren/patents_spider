# !/usr/bin/env python
# -*-coding: utf-8-*-
import xlsxwriter


# 生成excel文件
def generate_excel(rec_data,name):
    workbook = xlsxwriter.Workbook('./file/{}.xlsx'.format(name))
    worksheet = workbook.add_worksheet()

    bold_format = workbook.add_format({'bold': True})
     # 将二行二列设置宽度为15(从0开始)
    worksheet.set_column(1, 1, 15)

     # 用符号标记位置，例如：A列1行
    worksheet.write('A1', 'title', bold_format)
    worksheet.write('B1', 'url', bold_format)
    worksheet.write('C1', 'filing_date', bold_format)
    worksheet.write('D1', 'applicant', bold_format)
    worksheet.write('E1', 'main_classification_number', bold_format)
    worksheet.write('F1', 'inventor', bold_format)

    row = 1
    col = 0
    for item in (rec_data):
        # 使用write_string方法，指定数据格式写入数据
        worksheet.write_string(row, col, item['title'])
        worksheet.write_string(row, col + 1, item['url'])
        worksheet.write_string(row, col + 2, item['filing_date'])
        worksheet.write_string(row, col + 3, item['applicant'])
        worksheet.write_string(row, col + 4, item['main_classification_number'])
        worksheet.write_string(row, col + 5, item['inventor'])
        row += 1
    workbook.close()

    # coding=utf-8



# 读取execl
import xlrd


def read_xlrd(excelFile):
    data = xlrd.open_workbook(excelFile)
    table = data.sheet_by_index(1)
    dataFile = []

    for rowNum in range(table.nrows):
        # if 去掉表头
        if rowNum > 0:
            dataFile.append(table.row_values(rowNum))

    return dataFile


if __name__ == '__main__':
    excelFile = 'file/demo.xlsx'
    print(read_xlrd(excelFile=excelFile))