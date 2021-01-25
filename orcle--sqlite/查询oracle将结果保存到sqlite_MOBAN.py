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
A.TAG_INDEX,
A.TAG_NAME,
A.TAG_DESC,
A.DATE_TIME,
A.VALUE
FROM
IHTODB.TBL_IHTODB A
WHERE
A.DATE_TIME BETWEEN to_date('2020-01-01 00:00:00','YYYY-MM-DD HH24:MI:SS') AND to_date('2021-11-5 00:00:00','YYYY-MM-DD HH24:MI:SS') AND
to_char(date_time,'MI') = '00' AND
mod(to_number(to_char(date_time,'HH24')),24) = 0 
AND
A.TAG_INDEX = 134
ORDER BY
A.DATE_TIME ASC


"""


cursor.execute(sql3)
alldata_oracle = cursor.fetchall() 
# print(row)
print("========SUCCESS to  get  all data FROM ORACLE========")


# 
# print("====================查询结果写入sqlite")
import sqlite3
cn = sqlite3.connect(r'G:/98sqlite/外围气体/IH.db3')
# # G:\98sqlite\外围气体
#     # 设定表名称及结构
# # cn.execute('''CREATE TABLE IF NOT EXISTS FISRICASSEMBLE
# #     (
# #     PRODUCTNO TEXT PRIMARY KEY, 
# #     CYLINDER TEXT,
# #     MACHINE TEXT,
# #     PREFORM_DIAM TEXT,
# #     RIC_BEGIN DATE,
# #     RIC_END DATE,
# #     SUBTYPE TEXT,
# #     CONELENGTH INT,
# #     CYLINDERLENGTH INT,
# #     DRAWLENGTH INT ,
# #     RICNUM TEXT,
# #     SCRNUM TEXT,
# #     FIBRETYPE1 TEXT
# #     );''')
# print("===完成create table in sqlite===========")



for t in alldata_oracle:
    cn.execute("insert into IH values (?,?,?,?,?)", t)
cn.commit()  
cn.close()  
cursor.close
connection.close
print("===完成数据插入e===========")













 