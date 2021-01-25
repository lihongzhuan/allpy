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
    # cn.execute('''CREATE TABLE IF NOT EXISTS G652
    #     ( 
    #     FIBRE_NO TEXT PRIMARY KEY,
    #     DELIVERY_LENGTH  float,
    #     DIA_FIB_B float,
    #     DIA_BUF_B float,
    #     DIA_CLD_B float,
    #     CIR_CLD_B float,
    #     CON_BUF_B float,
    #     CON_FIB_B  float
    #     CON_COR_B float
    #     Curl   float
    #     ATT1310 float，
    #     ATT1383  float
    #     ATT1550 float,
    #     ATT1625float,
    #     MFD1310 float,
    #     CUT_OFF float,
    #     PMD_COEF float,
    #     DISP_ZERO_wavelentth float,
    #     DISP_SLOPE float,
    #     DISP1550 float,
    #     DISP1625 float,
    #     M301550 float,
    #     M301625 float,
    #     M201550 float,
    #     M201625 float,
    #     M101550 float,
    #     M101625 float,
    #     PREFORM_TYPES text，
    #     month text，
    #     F_Grade text，
    #     F_Type text，
    #     SMALL_TYPE text，
    #     ATT_MISL1310 float,
    #     ATT_MISL1383 float,
    #     ATT_MISL1550 float,
    #     TowerNo text);''')
    # print("===完成建表===")
    
    
    
    cn.execute('''CREATE TABLE IF NOT EXISTS kkk
        (
        PRODUCTNO TEXT ,  
        F_DeptNo TEXT,
        F_BatchNo TEXT,
        F_Status TEXT,
        F_FibreNo TEXT PRIMARY KEY,,
        F_ProdTime TEXT ,
        F_Oper TEXT,
        F_Code TEXT,
        F_Type TEXT,
        F_Grade TEXT,
        F_Color TEXT,
        F_Len INT,
        F_LenKm FLOAT,
        JobNo TEXT,
        ItemNo TEXT,
        Fibre_Status TEXT,
        Limitation TEXT,
        Dia_Buf_B FLOAT,
        Dia_Fib_B FLOAT，
        fn006_AS_Dia_Cor_B FLOAT,
        fn007_AS_Dia_Cld_B FLOAT,
        fn008_AS_Cir_Buf_B FLOAT,
        fn009_AS_Cir_Fib_B FLOAT,
        fn010_AS_Cir_Cor_B FLOAT,
        fn011_AS_Cir_Cld_B FLOAT,
        fn012_AS_Con_Buf_B FLOAT,
        fn013_AS_Con_Fib_B FLOAT,
        fn014_AS_Con_Cor_B FLOAT,
        fn015_AS_Bot_Reject TEXT,
        fn016_AS_Dia_Buf_T  FLOAT,
        fn017_AS_Dia_Fib_T FLOAT,
        fn018_AS_Dia_Cor_T FLOAT,
        fn019_AS_Dia_Cld_T FLOAT,
        fn020_AS_Cir_Buf_T FLOAT,
        fn021_AS_Cir_Fib_T FLOAT,
        fn022_AS_Cir_Cor_T FLOAT,
        fn023_AS_Cir_Cld_T FLOAT,
        fn024_AS_Con_Buf_T FLOAT,
        fn025_AS_Con_Fib_T FLOAT,
        fn026_AS_Con_Cor_T FLOAT,
        fn027_AS_Top_Reject TEXT,
        fn028_AS_Pmd400_EquipNo TEXT,
        fn029_AS_Pmd400_EquipTime TEXT,
        fn030_AS_Pmd400_Oper TEXT,
        fn031_AS_Pmd400_Pmd  FLOAT,
        fn032_AS_Pmd400_PmdCoef FLOAT,
        fn033_AS_Pmd400NT FLOAT,
        fn034_AS_HoldStore TEXT,
        fn039_AS_CD400_EquipNo TEXT,
        fn040_AS_CD400_EquipTime TEXT,
        fn041_AS_CD400_Oper TEXT,
        fn042_AS_Disp_1285 FLOAT,
        fn044_AS_Disp_1310 FLOAT,
        fn045_AS_Disp_1330 FLOAT,
        fn046_AS_Disp_1340 FLOAT,
        fn047_AS_Disp_1525 FLOAT,
        fn048_AS_Disp_1530 FLOAT,
        fn049_AS_Disp_1535 FLOAT,
        fn050_AS_Disp_1540 FLOAT,
        fn051_AS_Disp_1545 FLOAT,
        fn052_AS_Disp_1550 FLOAT,
        fn053_AS_Disp_1560 FLOAT,
        fn054_AS_Disp_1565 FLOAT,
        fn055_AS_Disp_1570 FLOAT,
        fn056_AS_Disp_1575 FLOAT,
        fn057_AS_Disp_1625 FLOAT,
        fn058_AS_Disp_Zero FLOAT,
        fn059_AS_Disp_Slope FLOAT,
        fn119_AS_Disp1550_slope FLOAT,
        fn121_AS_Disp1290 FLOAT,
        fn060_AS_Disp_Flag TEXT,
        fn061_AS_Disp_Reject TEXT,
        fn062_AS_RDS TEXT,
        fn063_AS_Total TEXT,
        fn064_AS_MFD1310B FLOAT,
        fn065_AS_MFD1550B FLOAT,
        fn043_AS_ATT1310MISL1 FLOAT,
        fn066_AS_ATT1383MISL FLOAT,
        fn067_AS_OTDR_LEN FLOAT,
        fn068_AS_MFD_Pig FLOAT,
        fn069_AS_Att1310 FLOAT,
        fn070_AS_D1310Dif FLOAT,
        fn071_AS_Dist1310 FLOAT,
        fn072_AS_AttUniF1310 FLOAT,
        fn073_AS_OTDRUniF1310 FLOAT,
        fn074_AS_OAtt1550 FLOAT,
        fn075_AS_D1550Dif FLOAT,
        fn076_AS_Dist1550 FLOAT,
        fn077_AS_AttUni1550 FLOAT,
        fn078_AS_OTDRUniF1550 FLOAT,
        fn079_AS_Tension_F FLOAT,
        fn082_AS_LOSST FLOAT,
        fn083_AS_LOSSB FLOAT,
        fn099_AS_ATT1550MISL FLOAT,
        fn141_AS_Cut_off FLOAT,
        fn142_AS_D2200MFD FLOAT,
        fn143_AS_MAC FLOAT,
        fn144_AS_Mac_Grade FLOAT,
        fn145_AS_Att_1285 FLOAT,
        fn146_AS_Att_1300 FLOAT,
        fn147_AS_Att_1310 FLOAT,
        fn148_AS_Att_1330 FLOAT,
        fn149_AS_Att_1340 FLOAT,
        fn150_AS_Att_1385 FLOAT,
        fn151_AS_Att_1383 FLOAT,
        fn152_AS_Att_1475 FLOAT,
        fn153_AS_Att_1525 FLOAT,
        fn154_AS_Att_1545 FLOAT,
        fn155_AS_Att_1565 FLOAT,
        fn156_AS_Att_1550 FLOAT,
        fn157_AS_Att_1575 FLOAT,
        fn158_AS_Att_1230 FLOAT,
        fn159_AS_Att_1240 FLOAT,
        fn160_AS_Att_1245 FLOAT,
        fn161_AS_Att_1250 FLOAT,
        fn162_AS_Att_1255 FLOAT,
        fn163_AS_Att_1260 FLOAT,
        fn164_AS_Att_1270 FLOAT,
        fn165_AS_Att_1280 FLOAT
        fn166_AS_Att_1290 FLOAT,
        fn167_AS_Att_1320 FLOAT,
        fn168_AS_Att_1335 FLOAT,
        fn169_AS_Att_1350 FLOAT,
        fn170_AS_Att_1360 FLOAT,
        fn171_AS_Att_1370 FLOAT,
        fn172_AS_Att_1380 FLOAT,
        fn173_AS_Att_1390 FLOAT,
        fn174_AS_Att_1400 FLOAT,
        fn175_AS_Att_1410 FLOAT,
        fn176_AS_Att_1420 FLOAT,
        fn177_AS_Att_1430 FLOAT,
        fn178_AS_Att_1440 FLOAT,
        fn179_AS_Att_1450 FLOAT,
        fn180_AS_Att_1460 FLOAT,
        fn181_AS_Att_1470 FLOAT,
        fn182_AS_Att_1480 FLOAT,
        fn183_AS_Att_1490 FLOAT,
        fn184_AS_Att_1500 FLOAT,
        fn185_AS_Att_1510 FLOAT,
        fn186_AS_Att_1520 FLOAT,
        fn187_AS_Att_1530 FLOAT,
        fn189_AS_Att_1540 FLOAT,
        fn190_AS_Att_1560 FLOAT,
        fn191_AS_Att_1570 FLOAT,
        fn192_AS_Att_1580 FLOAT,
        fn193_AS_Att_1590 FLOAT,
        fn194_AS_Att_1600 FLOAT,
        fn195_AS_Att_1620 FLOAT,
        fn196_AS_Att_1625 FLOAT,
        fn197_AS_Att_1650 FLOAT,
        fn201_AS_Lf_NY text,
        fn202_AS_Twist FLOAT,
        fn203_AS_Twist_Direct FLOAT,
        fn204_AS_Curl FLOAT,
        fn205_AS_TowerNo TEXT,
        fn188_AS_CCUTOFF FLOAT,
        fn096_AS_M10_1550 FLOAT,
        fn098_AS_M10_1625 FLOAT,
        fn092_AS_M30_1550 FLOAT,
        fn093_AS_M30_1625 FLOAT,
        Left([F_FibreNo],9)_AS_drum TEXT,
        Left([F_FibreNo],10)_AS_f10 TEXT,
        Left([F_FibreNo],8)_AS_rod TEXT,
        fn094_AS_M20_1550 FLOAT,
        fn097_AS_M20_1625 FLOAT,
        fn122_AS_AttMISL1310 FLOAT,
        fn123_AS_AttMISL1383 FLOAT,
        fn124_AS_AttMISL1550 FLOAT,
        Location TEXT,
        fn126_AS_[minAtt1310-1383] FLOAT,
        fn125_AS_week TEXT,
        fn120_AS_[盘具] TEXT,
        fn209_AS_Status TEXT,
        Left([F_FibreNo],10)_AS_fn10 TEXT,
        fn215_AS_D2_bn TEXT,
        fn036_AS_Preformtype TEXT,
        fn212 TEXT);''')
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