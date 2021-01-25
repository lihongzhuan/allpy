# -*- coding: utf-8 -*-
"""
Created on Tue Dec 22 20:44:21 2020

@author: 0100904
"""

import os
import platform
import xlsxwriter
import sqlite3
import time

##======================设置sqlite数据库相关信息+workbook
source = os.path.expanduser(r'g:\98sqlite\DRUM制造信息\DURM_DRAWING.db3')
cn =sqlite3.connect(source)
print('连接的数据库:',source)
curs = cn.cursor()

t = time.strftime('%Y.%m.%d',time.localtime(time.time()))
bookname = "d:/拉丝投入.xlsx"
book = xlsxwriter.Workbook(
    filename=bookname,
    options={  # 全局设置
        'strings_to_numbers': True,  # str 类型数字转换为 int 数字
        'strings_to_urls': False,  # 自动识别超链接
        'constant_memory': False,  # 连续内存模式 (True 适用于大数据量输出)
        'default_format_properties': {
            'font_name': '微软雅黑',  # 字体. 默认值 "Arial"
            'font_size': 9,  # 字号. 默认值 11
            # 'bold': False,  # 字体加粗
            # 'border': 1,  # 单元格边框宽度. 默认值 0
            'align': 'center',  # 对齐方式
            'valign': 'vcenter',  # 垂直对齐方式
            # 'text_wrap': False,  # 单元格内是否自动换行
            # ...
        },
    }
)
forma_head = book.add_format({'num_format':' ','font_size':'12','bold':'true'})
forma_date = book.add_format({'num_format':'yyyy-mm-dd','font_size':'9'})
forma_int = book.add_format({'num_format':'0','font_size':'9'})
forma_decimal = book.add_format({'num_format':'#,##.00','font_size':'9'})
forma_char = book.add_format({'num_format':' ','font_size':'9'})
print("写入的工作表为"+ bookname)





###sheet1START========================
sheet1name = '分塔'
sheet1 = book.add_worksheet(sheet1name)
sql_sheet1 = '''
SELECT
MACHINE,
sum(LENGTH) AS INPUT_KM,
MIN(DRAWING_END) as start,
MAX(DRAWING_END) as end
FROM  durm_drawing
WHERE
DRAWING_END > date('2020-01-01')
and
DRAWING_END < date('2020-12-31')
and 
MACHINE > 'T10'
AND
MACHINE < 'T53'
GROUP BY
MACHINE
 '''
curs.execute(sql_sheet1)
alldata_sqlite = curs.fetchall() 

fields = [field[0] for field in curs.description]  # 获取所有字段名    

# print(fields)
for col,field in enumerate(fields):
#     print(col,field)
    sheet1.write(0,col,field)
# print ("========success 完成写表头=========")    
row = 1
for data in alldata_sqlite:
    #print(data)
#     print ("d%",row) 
    for col,field in enumerate(data):
        sheet1.write(row,col,field)
        #print(type(row))
        #print（row）
       
        #print（field）
    print('写入的条数',row)
    row += 1
sheet1.set_column('A:Z',20)    
sheet1.set_row(0,{'font_size':'12','bold':'true'})

print("===完成写入并保存TABLE1：",sheet1name)
  
 #####shee1end============================================   
    
 




###sheet2 START========================
sheet2name = 'week'
sheet2 = book.add_worksheet(sheet2name)
# print(time.time('now'))
sql_sheet2 = '''
SELECT
strftime('%W',DRAWING_END) as weekno,
sum(LENGTH) AS INPUT_KM,
MIN(DRAWING_END) as start,
MAX(DRAWING_END) as end
FROM  durm_drawing
WHERE
DRAWING_END > date('2020-01-01')
and
DRAWING_END < date('2020-12-31')
and 
MACHINE > 'T10'
AND
MACHINE < 'T53'
GROUP BY
strftime('%W',DRAWING_END)
 '''
curs.execute(sql_sheet2)
alldata_sqlite = curs.fetchall() 

fields = [field[0] for field in curs.description]  # 获取所有字段名    

# print(fields)
for col,field in enumerate(fields):
#     print(col,field)
    sheet2.write(0,col,field)
# print ("========success 完成写表头=========")    
row = 1
for data in alldata_sqlite:
    #print(data)
#     print ("d%",row) 
    for col,field in enumerate(data):
        sheet2.write(row,col,field)
        #print(type(row))
        #print（row）
       
        #print（field）
    print('写入的条数',row)
    row += 1
sheet2.set_column('A:Z',20)    
sheet2.set_row(0,{'font_size':'12','bold':'true'})

print("===完成写入并保存TABLE1：",sheet2name)
  
 #####sheet2 end============================================   
    





###sheet3 START========================
sheet3name = 'subtype'
sheet3 = book.add_worksheet(sheet3name)
# print(time.time('now'))
sql_sheet3 = '''
SELECT
SUBTYPE,
sum(LENGTH) AS INPUT_KM,
MIN(DRAWING_END) as start,
MAX(DRAWING_END) as end
FROM  durm_drawing
WHERE
DRAWING_END > date('2020-01-01')
and
DRAWING_END < date('2020-12-31')
and 
MACHINE > 'T10'
AND
MACHINE < 'T53'
GROUP BY
SUBTYPE
 '''
curs.execute(sql_sheet3)
alldata_sqlite = curs.fetchall() 

fields = [field[0] for field in curs.description]  # 获取所有字段名    

# print(fields)
for col,field in enumerate(fields):
#     print(col,field)
    sheet3.write(0,col,field)
# print ("========success 完成写表头=========")    
row = 1
for data in alldata_sqlite:
    #print(data)
#     print ("d%",row) 
    for col,field in enumerate(data):
        sheet3.write(row,col,field)
        #print(type(row))
        #print（row）
       
        #print（field）
    print('写入的条数',row)
    row += 1
sheet3.set_column('A:Z',20)    
sheet3.set_row(0,{'font_size':'12','bold':'true'})

print("===完成写入并保存TABLE1：",sheet3name)
  
 #####sheet3 end============================================   





###sheet4 START========================
sheet4name = 'YEAR'
sheet4 = book.add_worksheet(sheet4name)
# print(time.time('now'))
sql_sheet4 = '''
SELECT
strftime('%Y',DRAWING_END) as YEARNO,
sum(LENGTH) AS INPUT_KM,
MIN(DRAWING_END) as start,
MAX(DRAWING_END) as end
FROM  durm_drawing
WHERE
MACHINE > 'T10'
AND
MACHINE < 'T53'
GROUP BY
strftime('%Y',DRAWING_END)
 '''
curs.execute(sql_sheet4)
alldata_sqlite = curs.fetchall() 

fields = [field[0] for field in curs.description]  # 获取所有字段名    

# print(fields)
for col,field in enumerate(fields):
#     print(col,field)
    sheet4.write(0,col,field)
# print ("========success 完成写表头=========")    
row = 1
for data in alldata_sqlite:
    #print(data)
#     print ("d%",row) 
    for col,field in enumerate(data):
        sheet4.write(row,col,field)
        #print(type(row))
        #print（row）
       
        #print（field）
    print('写入的条数',row)
    row += 1
sheet4.set_column('A:Z',20)    
sheet4.set_row(0,{'font_size':'12','bold':'true'})

print("===完成写入并保存TABLE1：",sheet4name)
  
 #####sheet4 end============================================   


















###sheet5 START========================
sheet5name = 'MONTH'
sheet5 = book.add_worksheet(sheet5name)
# print(time.time('now'))
sql_sheet5 = '''
SELECT
strftime('%m',DRAWING_END) as MONTHkno,
sum(LENGTH) AS INPUT_KM,
MIN(DRAWING_END) as start,
MAX(DRAWING_END) as end
FROM  durm_drawing
WHERE
DRAWING_END > date('2020-01-01')
and
DRAWING_END < date('2020-12-31')
and 
MACHINE > 'T10'
AND
MACHINE < 'T53'
GROUP BY
strftime('%m',DRAWING_END)
 '''
curs.execute(sql_sheet5)
alldata_sqlite = curs.fetchall() 

fields = [field[0] for field in curs.description]  # 获取所有字段名    

# print(fields)
for col,field in enumerate(fields):
#     print(col,field)
    sheet5.write(0,col,field)
# print ("========success 完成写表头=========")    
row = 1
for data in alldata_sqlite:
    #print(data)
#     print ("d%",row) 
    for col,field in enumerate(data):
        sheet5.write(row,col,field)
        #print(type(row))
        #print（row）
       
        #print（field）
    print('写入的条数',row)
    row += 1
sheet5.set_column('A:Z',20)    
sheet5.set_row(0,{'font_size':'12','bold':'true'})

print("===完成写入并保存TABLE1：",sheet5name)
  
 #####sheet5 end============================================   













  
###进度条
# import time
# scale = 50
# print("执行开始，祈祷不报错".center(scale // 2,"-"))
# start = time.perf_counter()
# for i in range(scale + 1):
#  a = "*" * i
#  b = "." * (scale - i)
#  c = (i / scale) * 100
#  dur = time.perf_counter() - start
#  print("\r{:^3.0f}%[{}->{}]{:.2f}s".format(c,a,b,dur),end = "")
#  time.sleep(0.1)
# print("\n"+"执行结束，万幸".center(scale // 2,"-"))
    
    
    
    
    
    
    

# sheet.set_column(1,None,forma)
# sheet.set_column("A:Z", 40)    #设置列宽度
# sheet.set_row(0, 30)    #设置行高度
# # format = book.add_format({'color':'red'})    #获取单元格属性
# #{'bold': True, 'font_size': 14, 'align': 'center','valign': 'vcenter','border':1, 'color':'red', 'bg_color':'blue'}
# dir(format)    #可以显示属性的种类
# # format.set_bold("A:A")    #设置粗体
# sheet.set_row(0,None,bold)
curs.close()   
# print("========完成1st表=========")
cn.commit()  
cn.close()  

cn.close()

book.close()
print("===完成全部===========")













 