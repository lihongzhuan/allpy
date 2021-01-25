# -*- coding: utf-8 -*-
"""
Created on Mon Dec 21 16:53:50 2020

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



## =====================链接数据库============
dsn = cx.makedsn('ora-dr1.yofc.com', 1521, 'orayofc')
connection = cx.connect('fis', 'fis', dsn)
cursor = connection.cursor() 
print("========SUCCESS connect to oracle=========")
## =======================开始1st表============================================
# sql1 = "*"
# sql2 = "select * from user_tables"
# cursor.execute(sql2)

sql1 = """
SELECT
to_char(max(FIS.FISDRAWINGPARA.DRAWING_END),'yyyy-mm-dd') as ends,
to_char(min(FIS.FISDRAWINGPARA.DRAWING_END),'yyyy-mm-dd') as starts,
round(sum(FIS.FISDRAWINGPARA.LENGTH),0) as 拉丝投入公里,
to_char(FIS.FISDRAWINGPARA.DRAWING_END,'mm') as monthid


FROM
FIS.FISDRAWINGPARA
WHERE
FIS.FISDRAWINGPARA.DRAWING_END > to_date('2020-1-1','yyyy-mm-dd')
and
FIS.FISDRAWINGPARA.DRAWING_END < to_date('2021-1-1','yyyy-mm-dd')
and 

fis.FISDRAWINGPARA.MACHINE > 'T10'
AND
fis.FISDRAWINGPARA.MACHINE < 'T53'
GROUP BY
to_char(FIS.FISDRAWINGPARA.DRAWING_END,'mm')
ORDER BY
to_char(FIS.FISDRAWINGPARA.DRAWING_END,'mm') DESC

 
 
"""

cursor.execute(sql1)
alldata = cursor.fetchall() 

# ====================查询结果写入excel
t = time.strftime('%Y.%m.%d',time.localtime(time.time()))
# print(t)

bookname = "拉丝产量" + t + ".xlsx"
print("写入的工作表为"+ bookname)
# ====================查询结果写入excel--月产量==============
book = xlsxwriter.Workbook(bookname)
sheet1 = book.add_worksheet('分月产量')
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
    sheet1.write(0,col,field)
# print ("========success 完成写表头=========")    
row = 1
for data in alldata:
    #print(data)
#     print ("d%",row) 
    for col,field in enumerate(data):
        sheet1.write(row,col,field)
        #print(type(row))
        #print（row）
        #print（col）
        #print（field）
    row += 1
#     print("========cesssfsff=========")

# sheet.set_column(1,None,forma)
# sheet.set_column("A:Z", 40)    #设置列宽度
# sheet.set_row(0, 30)    #设置行高度
# # format = book.add_format({'color':'red'})    #获取单元格属性
# #{'bold': True, 'font_size': 14, 'align': 'center','valign': 'vcenter','border':1, 'color':'red', 'bg_color':'blue'}
# dir(format)    #可以显示属性的种类
# # format.set_bold("A:A")    #设置粗体
# sheet.set_row(0,None,bold)
cursor.close()   
print("========完成1st表=========")


# =================
## =======================开始2nd表=============
# sql1 = "*"
# sql2 = "select * from user_tables"
# cursor.execute(sql2)

sql2 = """
SELECT
to_char(max(FIS.FISDRAWINGPARA.DRAWING_END),'yyyy-mm-dd') as ends,
to_char(min(FIS.FISDRAWINGPARA.DRAWING_END),'yyyy-mm-dd') as starts,
round(sum(FIS.FISDRAWINGPARA.LENGTH),0) as 拉丝投入公里,
FIS.FISDRAWINGPARA.MACHINE


FROM
FIS.FISDRAWINGPARA
WHERE
FIS.FISDRAWINGPARA.DRAWING_END > to_date('2020-1-1','yyyy-mm-dd')
and
FIS.FISDRAWINGPARA.DRAWING_END < to_date('2021-1-1','yyyy-mm-dd')
and 

fis.FISDRAWINGPARA.MACHINE > 'T10'
AND
fis.FISDRAWINGPARA.MACHINE < 'T53'
GROUP BY
FIS.FISDRAWINGPARA.MACHINE

ORDER BY
FIS.FISDRAWINGPARA.MACHINE desc


 
 
"""
cursor = connection.cursor() 
cursor.execute(sql2)
alldata = cursor.fetchall() 

# ====================查询结果写入excel
# t = time.strftime('%Y.%m.%d',time.localtime(time.time()))
# print(t)

# bookname = "拉丝产量" + t + ".xlsx"
# print(bookname)
# ====================查询结果写入excel--月产量==============
sheet2 = book.add_worksheet('分塔产量')
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
    sheet2.write(0,col,field)
# print ("========success 完成写表头=========")    
row = 1
for data in alldata:
    #print(data)
#     print ("d%",row) 
    for col,field in enumerate(data):
        sheet2.write(row,col,field)
        #print(type(row))
        #print（row）
        #print（col）
        #print（field）
    row += 1
#     print("========cesssfsff=========")

cursor.close()   

print("========完成2nd表=========")



# ================



# =======================开始3rd表=============
# sql1 = "*"
# sql2 = "select * from user_tables"
# cursor.execute(sql2)

sql3 = """
SELECT
to_char(max(FIS.FISDRAWINGPARA.DRAWING_END),'yyyy-mm-dd') as ends,
to_char(min(FIS.FISDRAWINGPARA.DRAWING_END),'yyyy-mm-dd') as starts,
round(sum(FIS.FISDRAWINGPARA.LENGTH),0) as 拉丝投入公里,
OUTPUTTYPE
FROM
FIS.FISDRAWINGPARA
WHERE
FIS.FISDRAWINGPARA.DRAWING_END > to_date('2020-1-1','yyyy-mm-dd')
and
FIS.FISDRAWINGPARA.DRAWING_END < to_date('2021-1-1','yyyy-mm-dd')
and 

fis.FISDRAWINGPARA.MACHINE > 'T10'
AND
fis.FISDRAWINGPARA.MACHINE < 'T53'
GROUP BY
OUTPUTTYPE
ORDER BY
OUTPUTTYPE DESC


 
 
"""
cursor = connection.cursor() 
cursor.execute(sql3)
alldata = cursor.fetchall() 

# ====================查询结果写入excel
# t = time.strftime('%Y.%m.%d',time.localtime(time.time()))
# print(t)

# bookname = "拉丝产量" + t + ".xlsx"
# print(bookname)
# ====================查询结果写入excel--月产量==============
sheet3 = book.add_worksheet('分棒型产量')
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
    sheet3.write(0,col,field)
# print ("========success 完成写表头=========")    
row = 1
for data in alldata:
    #print(data)
#     print ("d%",row) 
    for col,field in enumerate(data):
        sheet3.write(row,col,field)
        #print(type(row))
        #print（row）
        #print（col）
        #print（field）
    row += 1
#     print("========cesssfsff=========")

cursor.close()   

print("========完成3rd表=========")





# =======================开始4th表=============
# sql1 = "*"
# sql2 = "select * from user_tables"
# cursor.execute(sql2)

sql4 = """
SELECT
to_char(max(FIS.FISDRAWINGPARA.DRAWING_END),'yyyy-mm-dd') as ends,
to_char(min(FIS.FISDRAWINGPARA.DRAWING_END),'yyyy-mm-dd') as starts,
round(sum(FIS.FISDRAWINGPARA.LENGTH),0) as 拉丝投入公里,
to_char(FIS.FISDRAWINGPARA.DRAWING_END-2,'ww') as week_yofc


FROM
FIS.FISDRAWINGPARA
WHERE
FIS.FISDRAWINGPARA.DRAWING_END > to_date('2020-1-1','yyyy-mm-dd')
and
FIS.FISDRAWINGPARA.DRAWING_END < to_date('2021-1-1','yyyy-mm-dd')
and 

fis.FISDRAWINGPARA.MACHINE > 'T10'
AND
fis.FISDRAWINGPARA.MACHINE < 'T53'
GROUP BY
to_char(FIS.FISDRAWINGPARA.DRAWING_END-2,'ww') 

ORDER BY
to_char(FIS.FISDRAWINGPARA.DRAWING_END-2,'ww')






 
 
"""
cursor = connection.cursor() 
cursor.execute(sql4)
alldata = cursor.fetchall() 

# ====================查询结果写入excel
# t = time.strftime('%Y.%m.%d',time.localtime(time.time()))
# print(t)

# bookname = "拉丝产量" + t + ".xlsx"
# print(bookname)
# ====================查询结果写入excel--月产量==============
sheet4 = book.add_worksheet('分周产量')
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
    sheet4.write(0,col,field)
# print ("========success 完成写表头=========")    
row = 1
for data in alldata:
    #print(data)
#     print ("d%",row) 
    for col,field in enumerate(data):
        sheet4.write(row,col,field)
        #print(type(row))
        #print（row）
        #print（col）
        #print（field）
    row += 1
#     print("========cesssfsff=========")

cursor.close()   

print("========完成4th表=========")






# =======================开始5th表=============
# sql1 = "*"
# sql2 = "select * from user_tables"
# cursor.execute(sql2)

# sql5 = """
# SELECT
# to_char(max(FIS.FISDRAWINGPARA.DRAWING_END),'yyyy-mm-dd') as ends,
# to_char(min(FIS.FISDRAWINGPARA.DRAWING_END),'yyyy-mm-dd') as starts,
# round(sum(FIS.FISDRAWINGPARA.LENGTH),0) as 拉丝投入公里,
# OUTPUTTYPE
# FROM
# FIS.FISDRAWINGPARA
# WHERE
# FIS.FISDRAWINGPARA.DRAWING_END > to_date('2020-1-1','yyyy-mm-dd')
# and
# FIS.FISDRAWINGPARA.DRAWING_END < to_date('2021-1-1','yyyy-mm-dd')
# and 

# fis.FISDRAWINGPARA.MACHINE > 'T10'
# AND
# fis.FISDRAWINGPARA.MACHINE < 'T53'
# GROUP BY
# OUTPUTTYPE
# ORDER BY
# OUTPUTTYPE DESC


 
 
# """
# cursor = connection.cursor() 
# cursor.execute(sql5)
# alldata = cursor.fetchall() 

# # ====================查询结果写入excel
# # t = time.strftime('%Y.%m.%d',time.localtime(time.time()))
# # print(t)

# # bookname = "拉丝产量" + t + ".xlsx"
# # print(bookname)
# # ====================查询结果写入excel--月产量==============
# sheet5 = book.add_worksheet('分周产量')
#  # Add a bold format to use to highlight cells. 设置粗体，默认是False
# bold = book.add_format({'bold': True})
#  # Add a number format for cells with money.  定义数字格式
# money = book.add_format({'num_format': '$#,##0'})
#  # Add a number format for cells with money.  定义日期格式
# date_format = book.add_format({'num_format': 'd mmmm yyyy'})
# forma = book.add_format({'num_format':'yyyy-mm-dd'})
# fields = [field[0] for field in cursor.description]  # 获取所有字段名    
# # print(fields)
# for col,field in enumerate(fields):
# #     print(col,field)
#     sheet5.write(0,col,field)
# # print ("========success 完成写表头=========")    
# row = 1
# for data in alldata:
#     #print(data)
# #     print ("d%",row) 
#     for col,field in enumerate(data):
#         sheet5.write(row,col,field)
#         #print(type(row))
#         #print（row）
#         #print（col）
#         #print（field）
#     row += 1
# #     print("========cesssfsff=========")

# cursor.close()   

# print("========完成5th表=========")





book.close()
# print("========SUCCESS to excle=========")


  
connection.close()
print("========SUCCESS all！=========")



