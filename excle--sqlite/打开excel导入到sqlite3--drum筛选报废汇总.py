# -*- coding: utf-8 -*-
"""
Created on Tue Dec 22 16:11:08 2020

@author: 0100904
"""

#!/usr/bin/env python
# -*- coding:utf-8 -*-
# Author: Hogan


import xlrd
import sqlite3
import sqlite3

# def createDataBase(dbname):
#     cn = sqlite3.connect(dbname)
#     cn.execute('''CREATE TABLE IF NOT EXISTS TB_RIC
#     (
#     PRODUCTNO TEXT ,
#     CYLINDER TEXT,
#     MACHINE TEXT,
#     PREFORM_DIAM TEXT,
#     RIC_BEGIN DATE,
#     RIC_END DATE,
#     SUBTYPE TEXT,
#     CONELENGTH INT,
#     CYLINDERLENGTH INT,
#     DRAWLENGTH INT ,
#     RICNUM TEXT,
#     SCRNUM TEXT,
#     FIBRETYPE1 TEXT
#     );''')
#     print("===完成table1===")
#     cn.execute('''CREATE TABLE IF NOT EXISTS TB_RIT
#     (ID integer PRIMARY KEY AUTOINCREMENT,
#     PROVINCE TEXT,
#     TIME TEXT,
#     FILETYPE TEXT,
#     SCORE INTEGER);''')
#     print("===完成table2===")
# if __name__ == '__main__':
#     createDataBase('RIC+RIT组装记录集.db3')


def read_excel(fileName):
    # 打开文件excel
    workBook = xlrd.open_workbook(fileName)

    # 打开表格
    table = workBook.sheets()[0]
    # 计算文档有多少行
    all_row = table.nrows

    # 返回打开文档的对象，和文档的总行数
    return table, all_row


def create_con(dbnames):
    # conn = sqlite3.connect('example2.db')  # 连接数据库
    conn = sqlite3.connect(dbnames)  # 连接数据库
    # connect()方法，可以判断一个数据库文件是否存在，如果不存在就自动创建一个，如果存在的话，就打开那个数据库。
    cus = conn.cursor()  # 创建游标
    return cus, conn


def sql_dao(da, cus):
    # cus.execute('''CREATE TABLE stocks(id real ,lng REAL ,lat REAL,slp REAL ,intensity real ,utc text )''')

    # 向表中插入一条数据
    # sql2 = '''
    # insert into TB_RIC  values ('%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s'),(da[0],da[1], da[2], da[3], da[4],  da[5], da[6], da[7], da[8], da[9], da[10], da[11],da[12])

    tablename = 'sct_drum汇总'
    
    
    sql = '''INSERT INTO'''+ tablename + '''VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)''' % da
    cus.execute(sql)
    

def submit_close(conn):
    # 提交当前事务，保存数据
    conn.commit()
    # 关闭数据库连接
    conn.close()


if __name__ == '__main__':
    
    dbname = r'G:/98sqlite/筛选信息/sct.db3'
    xlsname = r'D:/users/查询8.xlsx'
    # createDataBase(dbname)
    row_all_obj, all_row_num = read_excel(xlsname)
    cus,conn = create_con(dbname)
    for i in range(1, all_row_num):
        data = row_all_obj.row_values(i)
        sql_dao(data,cus )
        print("插入第", i, "条")
        # conn.commit()
    # cus.close()        
    # conn.commit()
    # conn.close()        
    submit_close(conn)
    print("===========success===========")