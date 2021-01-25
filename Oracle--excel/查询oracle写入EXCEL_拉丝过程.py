# -*- coding: utf-8 -*-
"""
Created on Mon Dec 21 13:11:49 2020

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
#print(db.version)


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
FIS.FISEQUIPWORKINFO.ENDDATETIME,
FIS.FISEQUIPWORKINFO.EUIPNAME,
FIS.FISEQUIPWORKINFO.PROCESSNAME,
FIS.FISEQUIPWORKINFO.PRODUCTNO,
FIS.FISEQUIPWORKINFO.STARTDATETIME,
FIS.FISEQUIPWORKINFO.WORKCODE,
FIS.FISEQUIPWORKINFO.REMARK,
FIS.FISEQUIPWORKINFO.DRAWINGLEN,
FIS.FISEQUIPWORKINFO.WORKCODE2,
FIS.FISEQUIPWORKINFO.VERSION,
FIS.FISEQUIPWORKINFO.STATUS,
FIS.FISEQUIPWORKINFO.EQUIPMAINTAIN,
FIS.FISEQUIPWORKINFO.PROCESSCHECK,
FIS.FISEQUIPWORKINFO.PROCESSCONTOL,
FIS.FISEQUIPWORKINFO.FACTORY,
FIS.FISEQUIPWORKINFO.EDITOR,
FIS.FISEQUIPWORKINFO.REASON,
FIS.FISEQUIPWORKINFO.REJECTCODE,
FIS.FISEQUIPWORKINFO.WORKCODE3,
FIS.FISEQUIPWORKINFO.RESTARTTIMES,
FIS.FISEQUIPWORKINFO.EDITOR1,
FIS.FISEQUIPWORKINFO.TOWERBREAKREASON
FROM
FIS.FISEQUIPWORKINFO
WHERE
FIS.FISEQUIPWORKINFO.STARTDATETIME > to_date('2020-12-1 12:00:00','YYYY-MM-DD HH24:MI:SS') AND
FIS.FISEQUIPWORKINFO.EUIPNAME > 'T10' AND
FIS.FISEQUIPWORKINFO.EUIPNAME LIKE 'T%' AND
FIS.FISEQUIPWORKINFO.PRODUCTNO IS NOT NULL
ORDER BY
FIS.FISEQUIPWORKINFO.ENDDATETIME DESC


"""

## =======================执行SQL

cursor.execute(sql3)
alldata = cursor.fetchall() 
# print(row)
print("========SUCCESS to sql=========")


# ====================查询结果写入excel
t = time.strftime('%Y.%m.%d',time.localtime(time.time()))
# print(t)

bookname = "拉丝过程信息" + t + ".xlsx"
print("输出的文件名：" + bookname)

book = xlsxwriter.Workbook(bookname)
sheet = book.add_worksheet('sheet2')



book.set_custom_property('Checked by',       'LIHONGZHUAN')
book.set_custom_property('Date completed',   t)
book.set_custom_property('Document number',  12345)
book.set_custom_property('Reference number', 1.2345)
book.set_custom_property('Has review',       True)
book.set_custom_property('Signed off',       False)


forma = book.add_format({'num_format':'yyyy-mm-dd hh:mm:ss','font_size': 9, 'align': 'center','valign': 'vcenter'})
forma_head= book.add_format({'num_format':'#,##0.000','font_size': 10, 'align': 'center','valign': 'vcenter','bold':'true'})
forma_int = book.add_format({'num_format': '0','font_size': 9, 'align': 'center','valign': 'vcenter'})
forma_decimal = book.add_format({'num_format': '#,##0.0','font_size': 9, 'align': 'center','valign': 'vcenter'})
forma_date = book.add_format({'num_format':'yyyy-mm-dd hh:mm:ss','font_size': 9, 'align': 'center','valign': 'vcenter'})
forma_percentage = book.add_format({'num_format': '0.0%','font_size': 9, 'align': 'center','valign': 'vcenter'})
forma_char = book.add_format({'num_format': 's%','font_size': 9, 'align': 'center','valign': 'vcenter'})
forma_list = [forma_date,forma_char,forma_date,forma_date,forma_date,forma_char,forma,forma_int,forma,forma_int,forma,forma,forma,forma,forma,forma,forma,forma,forma,forma,forma,forma,forma,forma,forma,forma,forma,forma]






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
    sheet.write(0,col,field,forma_head)
print ("========success 完成写表头=========")    

row = 1
for data in alldata:
    #print(data)
#     print ("d%",row) 
    for col,field in enumerate(data):
        sheet.write(row,col,field,forma_list[col])
        #print(type(row))
        #print（row）
        #print（col）
        #print（field）
    row += 1
#     print("========cesssfsff=========")

# sheet.set_column(1,None,forma)
sheet.set_column("A:Z", 10)    #设置列宽度
sheet.set_column("G:G", 90)
sheet.set_column("E:E",20)
sheet.set_column("A:A",20)
sheet.set_row(0,10)    #设置行高度





# sheet.set_row(0, 18, date_format)
# sheet.set_column('A:A', 20,date_format)
# format = book.add_format({'color':'red'})    #获取单元格属性
#{'bold': True, 'font_size': 14, 'align': 'center','valign': 'vcenter','border':1, 'color':'red', 'bg_color':'blue'}
# dir(format)    #可以显示属性的种类
# format.set_bold("A:A")    #设置粗体
# 第一种
# cell_format = workbook.add_format()
# cell_format.set_bold()
# cell_format.set_font_color('red')
# 第二种
# cell_format = workbook.add_format({'bold': True, 'font_color': 'red'})
 
# 构造了Format对象并且已经设置了它的属性，它就可以作为参数传递给工作表write方法
# worksheet.write(0, 0, 'Foo', cell_format)
# worksheet.write_string(1, 0, 'Bar', cell_format)
# worksheet.write_number(2, 0, 3,     cell_format)
# worksheet.write_blank (3, 0, '',    cell_format)
 
# 也可以传递给工作表set_row()和set_column() 方法，以定义行或列的默认格式设置属性
# worksheet.set_row(0, 18, cell_format)
# worksheet.set_column('A:D', 20, cell_format)
sheet.set_row(0,None,bold)

# sheet.write(0, 0, 'Foo')
book.close()
# sheet.sort
print("========SUCCESS to excle=========")


cursor.close()   
connection.close()
print("========SUCCESS all！=========")



