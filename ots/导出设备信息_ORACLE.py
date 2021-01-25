# -*- coding: utf-8 -*-
"""
Created on Mon Dec 21 12:30:35 2020

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
FIS.FISEQUIPINFO.NAME,
FIS.FISEQUIPINFO.LOCATION,
FIS.FISEQUIPINFO.NAMEDESC,
FIS.FISEQUIPINFO.PROCSUBTYPE,
FIS.FISEQUIPINFO.PROCTYPE,
FIS.FISEQUIPINFO.STATUS,
FIS.FISEQUIPINFO.SEQ,
FIS.FISEQUIPINFO.RODTYPE,
FIS.FISEQUIPINFO.FIBRETYPE,
FIS.FISEQUIPINFO.REMARK,
FIS.FISEQUIPINFO.SPEED,
FIS.FISEQUIPINFO.STDCAPACITY,
FIS.FISEQUIPINFO.ABSDEVIATION,
FIS.FISEQUIPINFO.RELVALUE,
FIS.FISEQUIPINFO.COATINGTYPE,
FIS.FISEQUIPINFO.STDATE,
FIS.FISEQUIPINFO.STDATESEQ,
FIS.FISEQUIPINFO.FURNACESIZE,
FIS.FISEQUIPINFO.OUTPUTTYPE,
FIS.FISEQUIPINFO.RODLENGTHREQUIREMENT,
FIS.FISEQUIPINFO.THEORYSPEED
FROM
FIS.FISEQUIPINFO
WHERE
FIS.FISEQUIPINFO.PROCSUBTYPE = 'DRAWING'
ORDER BY
FIS.FISEQUIPINFO.NAME DESC



"""

## =======================执行SQL

cursor.execute(sql3)
alldata = cursor.fetchall() 
# print(row)
print("========SUCCESS to sql=========")


# ====================查询结果写入excel
t = time.strftime('%Y.%m.%d',time.localtime(time.time()))
print(t)

bookname = "设备信息汇总" + t + ".xlsx"
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


cursor.close()   
connection.close()
print("========SUCCESS all！=========")

