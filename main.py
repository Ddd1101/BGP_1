# -*- coding:utf-8 -*-

import xlrd
import os
import openpyxl
import xlwt
import re
import chardet
from datetime import datetime,date
import csv
import xlwt


if __name__ == "__main__":
    base_dir = os.getcwd()
    path = os.path.join(base_dir,'4.xlsx')
    workbook = xlrd.open_workbook(path)
    worksheet = workbook.sheet_by_name('Sheet 1')
    i = 1
    item = {}
    out_row = 0
    book = xlwt.Workbook()
    sheet = book.add_sheet('data')
    while i < worksheet.nrows-4:
        user_id = worksheet.cell_value(i, 1)
        j = 0
        while worksheet.cell_value(i, 1) == user_id:
            item[j] = worksheet.row(i)
            i += 1
        newitem = sorted(item.items(), key=lambda d: d[1][3], reverse=False)
        for k in newitem:
            out_col = 0
            if k[1][5].value == u'报告审核':
                k[1][5].value = u'报告发布'
            elif k[1][5].value == u'报告':
                k[1][5].value = u'报告发布'
            for h in k[1]:
                if h.ctype == 3:
                    tmp = xlrd.xldate_as_tuple(h.value, 0)
                    value = datetime(*tmp[:6]).strftime('%Y/%m/%d %H:%M:%S')
                    sheet.write(out_row, out_col, value)
                else:
                    sheet.write(out_row,out_col,h.value)
                out_col+=1
            out_row+=1
        book.save('out.xls')
