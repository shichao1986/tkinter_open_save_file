# coding:utf-8

import xlrd
import xlwt

import os

def excel_convert(filepath, output_name):
    curpath = os.path.dirname(__file__)
    filepath = os.path.join(curpath, filepath)
    excel_handle = xlrd.open_workbook(filepath)
    sheet0 = excel_handle.sheet_by_index(0)
    rows = sheet0.nrows
    excel_output = xlwt.Workbook(encoding='utf-8')
    excel_output_sheet = excel_output.add_sheet('MySheet1')
    for i in range(0, rows):
        line = sheet0.row_values(i)[0].encode('utf-8')
        for (j,val) in enumerate(line.split(',')):
            excel_output_sheet.write(i, j, label=val)
        print(line)

    target = os.path.join(curpath, output_name)
    # target.replace('\\','/')
    try:
        os.remove(target)
        excel_output.save(output_name)
    except Exception as e:
        print(e)


if __name__ == '__main__':
    excel_convert('kaoqin.xlsx', 'output.xls')
