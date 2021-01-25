# -*- coding: utf-8 -*-
'''
Created on Tue Dec 22 14:56:12 2020

@author: 0100904
'''
# /===============STEP1：/建立sqlite数据库======
import sqlite3

def createDataBase():
    cn = sqlite3.connect('check.db3')
    cn.execute('''CREATE TABLE IF NOT EXISTS TB_CHECK
    (ID integer PRIMARY KEY AUTOINCREMENT,
    PRODUCTNO TEXT,
    CYLINDER TEXT,
    MACHINE TEXT,
    PREFORM_DIAM TEXT,
    RIC_BEGIN DATE,
    RIC_END DATE,
    SUBTYPE TEXT,
    CONELENGTH INT,
    CYLINDERLENGTH INT,
    DRAWLENGTH INT ,
    RICNUM TEXT,
    SCRNUM TEXT,
    FIBRETYPE1 TEXT,
    TEMP TEXT);''')
    
    cn.execute('''CREATE TABLE IF NOT EXISTS TB_SCORE
    (ID integer PRIMARY KEY AUTOINCREMENT,
    PROVINCE TEXT,
    TIME TEXT,
    FILETYPE TEXT,
    SCORE INTEGER);''')
    print("===完成===")
if __name__ == '__main__':
    createDataBase()
# /===============STEP2：读取excel文件======
import xlrd
import xlwt
from datetime import date,datetime

def read_excel():
# 打开文件
import xlrd
import xlwt
from datetime import date,datetime
workbook = xlrd.open_workbook('d:/RIC组装表.xlsx')
# 获取所有sheet
sheet_name = workbook.sheet_names()[0]
sheet = workbook.sheet_by_name(sheet_name)

#获取一行的内容
for i in range(6,sheet.nrows):
for j in range(0,sheet.ncols):
print sheet.cell(i,j).value.encode('utf-8')

if __name__ == '__main__':
read_excel()
















