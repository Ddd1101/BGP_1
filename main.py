# -*- coding:utf-8 -*-

import xlrd
import os
import openpyxl
import xlwt

def fileload(path):
    workbook = xlrd.open_workbook(path)
    sheet_in = workbook.sheet_names()
    worksheet = workbook.sheet_by_name('Sheet 1')
    for i in range(1,worksheet.nrows):
        row = worksheet.row(i)
        for j in range(0,worksheet.ncols):
            print(worksheet.cell_value(i,j)," ")
        print('\n')


if __name__ == "__main__":
    base_dir = os.getcwd()
    path = os.path.join(base_dir,'4.xlsx')
    fileload(path)