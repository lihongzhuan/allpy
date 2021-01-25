# -*- coding: utf-8 -*-
"""
Created on Tue Dec 22 16:11:08 2020

@author: 0100904

不支持密文
"""

#!/usr/bin/env python
# -*- coding:utf-8 -*-
# Author: lihongzhuan


import xlrd
import sqlite3
import sqlite3

def createDataBase(dbname):
    cn = sqlite3.connect(dbname)
    cn.execute('''CREATE TABLE IF NOT EXISTS G652
        ( 
        FIBRE_NO TEXT PRIMARY KEY,
        DELIVERY_LENGTH  float,
        DIA_FIB_B float,
        DIA_BUF_B float,
        DIA_CLD_B float,
        CIR_CLD_B float,
        CON_BUF_B float,
        CON_FIB_B  float
        CON_COR_B float
        Curl   float
        ATT1310 float，
        ATT1383  float
        ATT1550 float,
        ATT1625float,
        MFD1310 float,
        CUT_OFF float,
        PMD_COEF float,
        DISP_ZERO_wavelentth float,
        DISP_SLOPE float,
        DISP1550 float,
        DISP1625 float,
        M301550 float,
        M301625 float,
        M201550 float,
        M201625 float,
        M101550 float,
        M101625 float,
        PREFORM_TYPES text，
        month text，
        F_Grade text，
        F_Type text，
        SMALL_TYPE text，
        ATT_MISL1310 float,
        ATT_MISL1383 float,
        ATT_MISL1550 float,
        TowerNo text);''')
    print("===完成建表===")
# cn.execute('''CREATE TABLE IF NOT EXISTS FISRICASSEMBLE
#     (
#     PRODUCTNO TEXT PRIMARY KEY, 
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
#     FIBRETYPE1 TEXT，
#     idcount INTEGER 
#     );''')
# print("===完成create table in sqlite===========")





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
    
    
    
    
 
    
    
    sql = ('INSERT INTO TB_RIC VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)',da)
    cus.execute(sql)
    

def submit_close(conn):
    # 提交当前事务，保存数据
    conn.commit()
    # 关闭数据库连接
    conn.close()


if __name__ == '__main__':
    
    dbname = r'G:/98sqlite/入库数据_PY_KPI.db3'
    xlsname = r'G:/98数据库/90DATA FORM PY/M2011/02 G.652 Fibre Monthly Quality Report (2001_2011).xlsx'
    createDataBase(dbname)
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