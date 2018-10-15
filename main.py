# -*- coding:utf-8 -*-

import xlrd
import os
import openpyxl

def fileload(path):
    workbook = openpyxl.load_workbook(path)
    #sheet_in = workbook.active
    sheet_in = workbook.get_sheet_by_name('Sheet 1')

    for row in sheet_in.rows:
        for cell in row:
            print(cell.value)
        print("\t")

    return sheet_in

def filewrite(path, sheet_in):
    workbook = openpyxl.load_workbook(path)
    sheet_out = workbook.active
    sheet_out.title = '1'

    for row in sheet_in:
        for cell in row:
            if cell.value == '':



if __name__ == "__main__":
    base_dir = os.getcwd()
    path = os.path.join(base_dir,'4.xlsx')
    fileload(path)