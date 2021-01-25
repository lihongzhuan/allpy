# -*- coding: utf-8 -*-
"""
Created on Tue Dec 22 20:44:21 2020

@author: 0100904
"""

import os
import platform
import xlsxwriter
import cx_Oracle as cx
import time

##======================设置oracle数据库相关信息
LOCATION = r"e:\oracle\client"
print("ARCH:", platform.architecture())
print("FILES AT LOCATION:")
for name in os.listdir(LOCATION):
    print(name)
os.environ["PATH"] = LOCATION + ";" + os.environ["PATH"]



## =====================链接数据库
dsn = cx.makedsn('ora-dr1.yofc.com', 1521, 'orayofc')
connection = cx.connect('fis', 'fis', dsn)
cursor = connection.cursor() 
print("========SUCCESS to oracle=========")
## =======================设定SQL
# sql1 = "*"
# sql2 = "select * from user_tables"
# cursor.execute(sql2)

sql3 = """
SELECT
FIS.FISRICASSEMBLEPARA.PRODUCTNO,
FIS.FISRICASSEMBLEPARA.CYLINDER,
FIS.FISRICASSEMBLEPARA.MACHINE,
FIS.FISRICASSEMBLEPARA.PREFORM_DIAM,
FIS.FISRICASSEMBLEPARA.RIC_BEGIN,
FIS.FISRICASSEMBLEPARA.RIC_END,
FIS.FISRICASSEMBLEPARA.SUBTYPE,
FIS.FISRICASSEMBLEPARA.CONELENGTH,
FIS.FISRICASSEMBLEPARA.CYLINDERLENGTH,
FIS.FISRICASSEMBLEPARA.DRAWLENGTH,
FIS.FISRICASSEMBLEPARA.RICNUM,
FIS.FISRICASSEMBLEPARA.SCRNUM,
FIS.FISRICASSEMBLEPARA.FIBRETYPE1
FROM
FIS.FISRICASSEMBLEPARA
WHERE
FIS.FISRICASSEMBLEPARA.RIC_BEGIN > to_date('2010-1-1 12:00:00','YYYY-MM-DD HH24:MI:SS')
ORDER BY
FIS.FISRICASSEMBLEPARA.RIC_BEGIN DESC


"""

## =======================执行SQL

cursor.execute(sql3)
alldata = cursor.fetchall() 
# print(row)
print("========SUCCESS to  get  all data FROM ORACLE========")


# 
# print("====================查询结果写入sqlite")
import sqlite3
cn = sqlite3.connect(r'd:/FISRICASSEMBLE2.db3')
    # 设定表名称及结构
cn.execute('''CREATE TABLE IF NOT EXISTS FISRICASSEMBLE
    (
    PRODUCTNO TEXT PRIMARY KEY, 
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
    FIBRETYPE1 TEXT
    );''')
print("===完成create table in sqlite===========")
    # print(sql)
    # cn.execute(sql)
    # cn = sqlite3.connect(dbname)
    # cn.execute('''CREATE TABLE IF NOT EXISTS TB_RIC
    # (
    # PRODUCTNO TEXT PRIMARY KEY, 
    # CYLINDER TEXT,
    # MACHINE TEXT,
    # PREFORM_DIAM TEXT,
    # RIC_BEGIN DATE,
    # RIC_END DATE,
    # SUBTYPE TEXT,
    # CONELENGTH INT,
    # CYLINDERLENGTH INT,
    # DRAWLENGTH INT ,
    # RICNUM TEXT,
    # SCRNUM TEXT,
    # FIBRETYPE1 TEXT
    # );''')
    # print("===完成table1===")
# print("===完写入在数据库sqlite")