# -*- coding: utf-8 -*-
"""
Created on Mon Dec 21 13:10:31 2020

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
# print(db.version)


## =====================链接数据库
dsn = cx.makedsn('ora-dr1.yofc.com', 1521, 'orayofc')
connection = cx.connect('fis', 'fis', dsn)
cursor = connection.cursor() 
print("========SUCCESS to oracle=========")
## =======================设定SQL
sql1 = "*"
sql2 = "select * from user_tables"
# cursor.execute(sql2)

sql3 = """
SELECT
FIS.FISDRAWINGRETURNPARA.CREATEDATETIME,
FIS.FISDRAWINGRETURNPARA.RETURNREASON,
-- FIS.FISDRAWINGRETURNPARA.CYLINDERNO,
-- FIS.FISDRAWINGRETURNPARA.PRODUCTNO,
FIS.FISDRAWINGRETURNPARA.TOEWR,
FIS.FISDRAWINGRETURNPARA.TYPE,
FIS.FISDRAWINGRETURNPARA.SUBTYPE,

FIS.FISDRAWINGRETURNPARA.CALDRAWLENGTH,
FIS.FISDRAWINGRETURNPARA.CUSTDRUMNO,
FIS.FISDRAWINGRETURNPARA.CYLINDERNO,
FIS.FISDRAWINGRETURNPARA.DIFFERLENGTH,
FIS.FISDRAWINGRETURNPARA.DRAWEDLENGTH,
FIS.FISDRAWINGRETURNPARA.PROCESSOPER,
FIS.FISDRAWINGRETURNPARA.PROCESSTIME,
FIS.FISDRAWINGRETURNPARA.PROCESSNY,
FIS.FISDRAWINGRETURNPARA.RETURNREASON,
FIS.FISDRAWINGRETURNPARA.RETURNREMARK,




FIS.FISDRAWINGRETURNPARA.CREATEDPROCESSNY,
FIS.FISDRAWINGRETURNPARA.CREATEDRETURNREASON,
FIS.FISDRAWINGRETURNPARA.ITEM

FROM
FIS.FISDRAWINGRETURNPARA
WHERE
FIS.FISDRAWINGRETURNPARA.PROCESSNY NOT LIKE '不需处理' AND

FIS.FISDRAWINGRETURNPARA.CREATEDATETIME BETWEEN to_date('2020-10-1 00:00:00','YYYY-MM-DD HH24:MI:SS') AND to_date('2021-1-1 00:00:00','YYYY-MM-DD HH24:MI:SS')
ORDER BY
FIS.FISDRAWINGRETURNPARA.CREATEDATETIME DESC

"""

## =======================执行SQL

cursor.execute(sql3)
alldata = cursor.fetchall() 
# print(row)
print("========SUCCESS to sql=========")


# ====================查询结果写入excel
t = time.strftime('%Y.%m.%d',time.localtime(time.time()))
print(t)

bookname = "退棒" + t + ".xlsx"
print(bookname)

book = xlsxwriter.Workbook(bookname)
sheet = book.add_worksheet('sheet2')
 # Add a bold format to use to highlight cells. 设置粗体，默认是False
bold = book.add_format({'bold': True})
 # Add a number format for cells with money.  定义数字格式
money = book.add_format({'num_format': '$#,##0'})
 # Add a number format for cells with money.  定义日期格式
date_format = book.add_format({'num_format': 'd mmmm yyyy'})
forma = book.add_format({'num_format':'yyyy-mm-dd'})


fields = [field[0] for field in cursor.description]  # 获取所有字段名    
# print(fields)
for col,field in enumerate(fields):
#     print(col,field)
    sheet.write(0,col,field)
print ("========success 完成写表头=========")    

row = 1
for data in alldata:
    #print(data)
#     print ("d%",row) 
    for col,field in enumerate(data):
        sheet.write(row,col,field)
        #print(type(row))
        #print（row）
        #print（col）
        #print（field）
    row += 1
#     print("========cesssfsff=========")

# sheet.set_column(1,None,forma)
sheet.set_column("A:Z", 40)    #设置列宽度
sheet.set_row(0, 30)    #设置行高度
# format = book.add_format({'color':'red'})    #获取单元格属性
#{'bold': True, 'font_size': 14, 'align': 'center','valign': 'vcenter','border':1, 'color':'red', 'bg_color':'blue'}
dir(format)    #可以显示属性的种类
# format.set_bold("A:A")    #设置粗体
sheet.set_row(0,None,bold)
book.close()
print("========SUCCESS to excle=========")


## ========================写入sqlite==================
# import sqlite3
# cx = sqlite3.connect("g:/退棒.db")
# sql_insert = '''
#     INSERT INTO
#       users(username, password, email)
#     VALUES
#       (?, ?, ?);
#     '''


print("========SUCCESS sqlite！=========")
cursor.close()   
connection.close()
print("========SUCCESS all！=========")



