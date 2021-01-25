# -*- coding: utf-8 -*-
"""
Created on Tue Dec 22 15:25:34 2020

@author: 0100904
不能读取加密文件

"""

import xlrd
import xlwt
from datetime import date,datetime

def read_excel():
# 打开文件
    workbook = xlrd.open_workbook(r'd:/RIC组装表2020.xlsx')
# 获取所有sheet
    sheet_name = workbook.sheet_names()[0]
    sheet = workbook.sheet_by_name(sheet_name)

#显示表内容
    for i in range(6,sheet.nrows):
            for j in range(0,sheet.ncols):
                print(sheet.cell(i,j).value)

if __name__ == '__main__':
    read_excel()