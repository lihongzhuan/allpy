# -*- coding: utf-8 -*-
"""
Created on Mon Dec 21 13:13:04 2020

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
# sql1 = "*"
sql2 = """
select 
FIS.FISDRAWINGPARA.PRODUCTNO,
FIS.FISDRAWINGPARA.LENGTH, 
FIS.FISDRAWINGPARA.DRAWING_END，
FIS.FISDRAWINGPARA.MACHINE，

FIS.FISDRAWINGPARA.POINTER, 

FIS.FISDRAWINGPARA.STATUS, 
FIS.FISDRAWINGPARA.SPEEDDOWNRESON,

FIS.FISDRAWINGPARA.PISTONNO, 
FIS.FISDRAWINGPARA.DRUMSN,
FIS.FISDRAWINGPARA.COUNTER,
FIS.FISDRAWINGPARA.TOWERBREAKTYPE,
FIS.FISDRAWINGPARA.PISTONNOCHANGE AS 更换了涂覆器, 
FIS.FISDRAWINGPARA.COATINGTYPE,
FIS.FISDRAWINGPARA.FIBRETYPE1, 
FIS.FISDRAWINGPARA.REMARK






FROM
FIS.FISDRAWINGPARA
WHERE
FIS.FISDRAWINGPARA.DRAWING_END > to_date('2020-12-19 12:00:00','YYYY-MM-DD HH24:MI:SS') AND
FIS.FISDRAWINGPARA.MACHINE > 'T10'  and
FIS.FISDRAWINGPARA.LENGTH > 2
order by
FIS.FISDRAWINGPARA.MACHINE desc，
FIS.FISDRAWINGPARA.DRAWING_END desc

"""

# 常用字段= 

# SELECT  
# FIS.FISDRAWINGPARA
# Round(FIS.FISDRAWINGPARA.LENGTH*1000/((FIS.FISDRAWINGPARA.DRAWING_END-FIS.FISDRAWINGPARA.DRAWING_BEGIN)*60*24),0) AS 拉丝速度,
# FIS.FISDRAWINGPARA.PRODUCTNO
# AAA_FISDRAWINGPARA.LENGTH, 
# AAA_FISDRAWINGPARA.DRAWING_END,
# AAA_FISDRAWINGPARA.POINTER, 
# # Round(1000*Sum([COUNTER])/Sum([POINTER]),0) AS 千公里断点数, 
# AAA_FISDRAWINGPARA.STATUS, 
# AAA_FISDRAWINGPARA.SPEEDDOWNRESON,
# # Round([LENGTH]*1000/(([DRAWING_END]-[DRAWING_BEGIN])*60*24),0) AS 拉丝速度,
# AAA_FISDRAWINGPARA.PISTONNO, 
# AAA_FISDRAWINGPARA.DRUMSN,
# AAA_FISDRAWINGPARA.COUNTER,
# AAA_FISDRAWINGPARA.TOWERBREAKTYPE,
# # AAA_FISDRAWINGPARA.PISTONNOCHANGE AS 更换了涂覆器, 
# AAA_FISDRAWINGPARA.COATINGTYPE,
# AAA_FISDRAWINGPARA.FIBRETYPE1, 
# AAA_FISDRAWINGPARA.REMARK
# FROM
# FIS.FISDRAWINGPARA
# # GROUP BY
# MACHINE,
# # AAA_FISDRAWINGPARA.PRODUCTNO,
# AAA_FISDRAWINGPARA.LENGTH, 
# AAA_FISDRAWINGPARA.DRAWING_END, 
# AAA_FISDRAWINGPARA.POINTER, 
# AAA_FISDRAWINGPARA.STATUS,
# # AAA_FISDRAWINGPARA.SPEEDDOWNRESON, Round([LENGTH]*1000/(([DRAWING_END]-[DRAWING_BEGIN])*60*24),0),
# AAA_FISDRAWINGPARA.PISTONNO,
# AAA_FISDRAWINGPARA.DRUMSN, 
# AAA_FISDRAWINGPARA.COUNTER, 
# AAA_FISDRAWINGPARA.TOWERBREAKTYPE, 
# AAA_FISDRAWINGPARA.PISTONNOCHANGE, 
# AAA_FISDRAWINGPARA.COATINGTYPE,
# AAA_FISDRAWINGPARA.FIBRETYPE1,
# AAA_FISDRAWINGPARA.REMARK, 
# AAA_FISDRAWINGPARA.DRAWING_END
# HAVING 
# FIS.FISDRAWINGPARA.MACHINE > 'T10'
# AND
# FIS.FISDRAWINGPARA.DRAWING_END >= ow()-1
# # # ORDER BY 
# # MACHINE, 
# # DRAWING_END;




## =======================执行SQL

cursor.execute(sql2)
alldata = cursor.fetchall() 
# print(row)
print("========SUCCESS to sql=========")


# ====================查询结果写入excel
t = time.strftime('%Y.%m.%d',time.localtime(time.time()))
# print(t)

bookname = "DRUM控制图" + t + ".xlsx"
print("输出的文件名：" + bookname)

book = xlsxwriter.Workbook(bookname)
sheet = book.add_worksheet('sheet2')



book.set_custom_property('Checked by',       'Eve')
book.set_custom_property('Date completed',   t)
book.set_custom_property('Document number',  12345)
book.set_custom_property('Reference number', 1.2345)
book.set_custom_property('Has review',       True)
book.set_custom_property('Signed off',       False)





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
sheet.set_column("A:Z", 20)    #设置列宽度
sheet.set_row(0,10)    #设置行高度





# sheet.set_row(0, 18, date_format)
sheet.set_column('A:A', 20,date_format)
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



