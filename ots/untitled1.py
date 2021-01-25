# -*- coding: utf-8 -*-
"""
Created on Tue Dec 22 15:25:34 2020

@author: 0100904
"""

import xlrd
import xlwt
from datetime import date,datetime
workbook = xlrd.open_workbook('d:\RIC组装表.xlsx')
# 获取所有sheet
sheet_name = workbook.sheet_names()[0]
sheet = workbook.sheet_by_name(sheet_name)

#获取一行的内容
for i in range(6,sheet.nrows):
    for j in range(0,sheet.ncols):
        print sheet.cell(i,j).value.encode('utf-8')