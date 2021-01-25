# -*- coding: utf-8 -*-
"""
Created on Wed Dec 23 10:56:32 2020

@author: 0100904
"""

import xlsxwriter
import datetime

workbook = xlsxwriter.Workbook("test.xlsx")
worksheet = workbook.add_worksheet("sheet1")
forma = workbook.add_format({'num_format':'yyyy-mm-dd','font_size': 9, 'align': 'center','valign': 'vcenter'})
head = [u'菜单路径']

worksheet.write_row('A1', head,forma)
print(type('a1'))

dat = ['2018.01.02', '2017.9.2', '2018.9.3', '2017.9.4', '2017.9.5', '2017.9.6']

des = []
for i in dat:
    dae = datetime.datetime.strptime(i,'%Y.%m.%d')
    print(type(dae))
    des.append(dae)
print(des)
forma = workbook.add_format({'num_format':'yyyy-mm-dd','font_size': 9})

worksheet.write_column('A2',des,forma)
worksheet.set_column("A:A",10)

# worksheet.set_column("A:A", 40)    #设置列宽度
#  7 worksheet.set_row(0, 30)    #设置行高度
#  8 
#  9 format = workbook.add_format({'color':'red'})    #获取单元格属性
# format = workbook.add_format({'color':'red'})    #获取单元格属性
# 10 #{'bold': True, 'font_size': 14, 'align': 'center','valign': 'vcenter','border':1, 'color':'red', 'bg_color':'blue'}
# 11 #dir(format)    #可以显示属性的种类
# 12 #format.set_bold()    #设置粗体


workbook.close()