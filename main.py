# -*- coding:utf-8 -*-

import xlrd
import os
import openpyxl

def fileload(path):
    workbook = openpyxl.load_workbook(path)
    sheet = workbook.active
    sheet = workbook.get_active_sheet()

    for row in sheet.rows:
        for cell in row:
            print(cell.value)
        print("\t")


if __name__ == "__main__":
    base_dir = os.getcwd()
    path = os.path.join(base_dir,'4.xlsx')
    fileload(path)