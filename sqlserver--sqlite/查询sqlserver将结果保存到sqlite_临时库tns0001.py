# -*- coding: utf-8 -*-
"""
Created on Tue Dec 22 20:44:21 2020

@author: 0100904
"""

import pymssql #引入pymssql模块
import xlwt
import xlsxwriter


connect = pymssql.connect('10.42.90.92', 'fis', 'fis', 'cab') #服务器名,账户,密码,数据库名

print("==========连接sqlserver成功!==========")
crsr = connect.cursor() 
# select name from sysobjects where xtype='u'
# select * from sys.tables


#查询全部表名称
# cursor = connect.cursor()   #创建一个游标对象,python里的sql语句都要通过cursor来执行
# sql = "select * from sys.tables"
# row = crsr.fetchone()  #读取查询结果,
# while row:              #循环读取所有结果
#     print("Name=%s, Sex=%s" % (row[0],row[1]))   #输出结果
   
#=================查询curl=====================

 
sql5 = """
SELECT 
dbo.tnS_S0001.F_DeptNo,
dbo.tnS_S0001.F_BatchNo,
dbo.tnS_S0001.F_Status,
dbo.tnS_S0001.F_FibreNo,
dbo.tnS_S0001.F_ProdTime,
dbo.tnS_S0001.F_Oper,
dbo.tnS_S0001.F_Code,
dbo.tnS_S0001.F_Type,
dbo.tnS_S0001.F_Grade,
dbo.tnS_S0001.F_Color,
dbo.tnS_S0001.F_Len,
dbo.tnS_S0001.F_LenKm,
dbo.tnS_S0001.JobNo,
dbo.tnS_S0001.ItemNo,
dbo.tnS_S0001.Fibre_Status,
dbo.tnS_S0001.Limitation,
dbo.tnS_S0001.fn004 AS Dia_Buf_B,
dbo.tnS_S0001.fn005 AS Dia_Fib_B,
dbo.tnS_S0001.fn006 AS Dia_Cor_B,
dbo.tnS_S0001.fn007 AS Dia_Cld_B,
dbo.tnS_S0001.fn008 AS Cir_Buf_B,
dbo.tnS_S0001.fn009 AS Cir_Fib_B,
dbo.tnS_S0001.fn010 AS Cir_Cor_B,
dbo.tnS_S0001.fn011 AS Cir_Cld_B,
dbo.tnS_S0001.fn012 AS Con_Buf_B,
dbo.tnS_S0001.fn013 AS Con_Fib_B,
dbo.tnS_S0001.fn014 AS Con_Cor_B,
dbo.tnS_S0001.fn015 AS Bot_Reject,
dbo.tnS_S0001.fn016 AS Dia_Buf_T,
dbo.tnS_S0001.fn017 AS Dia_Fib_T,
dbo.tnS_S0001.fn018 AS Dia_Cor_T,
dbo.tnS_S0001.fn019 AS Dia_Cld_T,
dbo.tnS_S0001.fn020 AS Cir_Buf_T,
dbo.tnS_S0001.fn021 AS Cir_Fib_T,
dbo.tnS_S0001.fn022 AS Cir_Cor_T,
dbo.tnS_S0001.fn023 AS Cir_Cld_T,
dbo.tnS_S0001.fn024 AS Con_Buf_T,
dbo.tnS_S0001.fn025 AS Con_Fib_T,
dbo.tnS_S0001.fn026 AS Con_Cor_T,
dbo.tnS_S0001.fn027 AS Top_Reject,
dbo.tnS_S0001.fn028 AS Pmd400_EquipNo,
dbo.tnS_S0001.fn029 AS Pmd400_EquipTime,
dbo.tnS_S0001.fn030 AS Pmd400_Oper,
dbo.tnS_S0001.fn031 AS Pmd400_Pmd,
dbo.tnS_S0001.fn032 AS Pmd400_PmdCoef,
dbo.tnS_S0001.fn033 AS Pmd400NT,
dbo.tnS_S0001.fn034 AS HoldStore,
dbo.tnS_S0001.fn039 AS CD400_EquipNo,
dbo.tnS_S0001.fn040 AS CD400_EquipTime,
dbo.tnS_S0001.fn041 AS CD400_Oper,
dbo.tnS_S0001.fn042 AS Disp_1285,
dbo.tnS_S0001.fn044 AS Disp_1310,
dbo.tnS_S0001.fn045 AS Disp_1330,
dbo.tnS_S0001.fn046 AS Disp_1340,
dbo.tnS_S0001.fn047 AS Disp_1525,
dbo.tnS_S0001.fn048 AS Disp_1530,
dbo.tnS_S0001.fn049 AS Disp_1535,
dbo.tnS_S0001.fn050 AS Disp_1540,
dbo.tnS_S0001.fn051 AS Disp_1545,
dbo.tnS_S0001.fn052 AS Disp_1550,
dbo.tnS_S0001.fn053 AS Disp_1560,
dbo.tnS_S0001.fn054 AS Disp_1565,
dbo.tnS_S0001.fn055 AS Disp_1570,
dbo.tnS_S0001.fn056 AS Disp_1575,
dbo.tnS_S0001.fn057 AS Disp_1625,
dbo.tnS_S0001.fn058 AS Disp_Zero,
dbo.tnS_S0001.fn059 AS Disp_Slope,
dbo.tnS_S0001.fn119 AS Disp1550_slope,
dbo.tnS_S0001.fn121 AS Disp1290,
dbo.tnS_S0001.fn060 AS Disp_Flag,
dbo.tnS_S0001.fn061 AS Disp_Reject,
dbo.tnS_S0001.fn062 AS RDS,
dbo.tnS_S0001.fn063 AS Total,
dbo.tnS_S0001.fn064 AS MFD1310B,
dbo.tnS_S0001.fn065 AS MFD1550B,
dbo.tnS_S0001.fn043 AS ATT1310MISL1,
dbo.tnS_S0001.fn066 AS ATT1383MISL,
dbo.tnS_S0001.fn067 AS OTDR_LEN,
dbo.tnS_S0001.fn068 AS MFD_Pig,
dbo.tnS_S0001.fn069 AS Att1310,
dbo.tnS_S0001.fn070 AS D1310Dif,
dbo.tnS_S0001.fn071 AS Dist1310,
dbo.tnS_S0001.fn072 AS AttUniF1310,
dbo.tnS_S0001.fn073 AS OTDRUniF1310,
dbo.tnS_S0001.fn074 AS OAtt1550,
dbo.tnS_S0001.fn075 AS D1550Dif,
dbo.tnS_S0001.fn076 AS Dist1550,
dbo.tnS_S0001.fn077 AS AttUni1550,
dbo.tnS_S0001.fn078 AS OTDRUniF1550,
dbo.tnS_S0001.fn079 AS Tension_F,
dbo.tnS_S0001.fn082 AS LOSST,
dbo.tnS_S0001.fn083 AS LOSSB,
dbo.tnS_S0001.fn099 AS ATT1550MISL,
dbo.tnS_S0001.fn141 AS Cut_off,
dbo.tnS_S0001.fn142 AS D2200MFD,
dbo.tnS_S0001.fn143 AS MAC,
dbo.tnS_S0001.fn144 AS Mac_Grade,
dbo.tnS_S0001.fn145 AS Att_1285,
dbo.tnS_S0001.fn146 AS Att_1300,
dbo.tnS_S0001.fn147 AS Att_1310,
dbo.tnS_S0001.fn148 AS Att_1330,
dbo.tnS_S0001.fn149 AS Att_1340,
dbo.tnS_S0001.fn150 AS Att_1385,
dbo.tnS_S0001.fn151 AS Att_1383,
dbo.tnS_S0001.fn152 AS Att_1475,
dbo.tnS_S0001.fn153 AS Att_1525,
dbo.tnS_S0001.fn154 AS Att_1545,
dbo.tnS_S0001.fn155 AS Att_1565,
dbo.tnS_S0001.fn156 AS Att_1550,
dbo.tnS_S0001.fn157 AS Att_1575,
dbo.tnS_S0001.fn158 AS Att_1230,
dbo.tnS_S0001.fn159 AS Att_1240,
dbo.tnS_S0001.fn160 AS Att_1245,
dbo.tnS_S0001.fn161 AS Att_1250,
dbo.tnS_S0001.fn162 AS Att_1255,
dbo.tnS_S0001.fn163 AS Att_1260,
dbo.tnS_S0001.fn164 AS Att_1270,
dbo.tnS_S0001.fn165 AS Att_1280,
dbo.tnS_S0001.fn166 AS Att_1290,
dbo.tnS_S0001.fn167 AS Att_1320,
dbo.tnS_S0001.fn168 AS Att_1335,
dbo.tnS_S0001.fn169 AS Att_1350,
dbo.tnS_S0001.fn170 AS Att_1360,
dbo.tnS_S0001.fn171 AS Att_1370,
dbo.tnS_S0001.fn172 AS Att_1380,
dbo.tnS_S0001.fn173 AS Att_1390,
dbo.tnS_S0001.fn174 AS Att_1400,
dbo.tnS_S0001.fn175 AS Att_1410,
dbo.tnS_S0001.fn176 AS Att_1420,
dbo.tnS_S0001.fn177 AS Att_1430,
dbo.tnS_S0001.fn178 AS Att_1440,
dbo.tnS_S0001.fn179 AS Att_1450,
dbo.tnS_S0001.fn180 AS Att_1460,
dbo.tnS_S0001.fn181 AS Att_1470,
dbo.tnS_S0001.fn182 AS Att_1480,
dbo.tnS_S0001.fn183 AS Att_1490,
dbo.tnS_S0001.fn184 AS Att_1500,
dbo.tnS_S0001.fn185 AS Att_1510,
dbo.tnS_S0001.fn186 AS Att_1520,
dbo.tnS_S0001.fn187 AS Att_1530,
dbo.tnS_S0001.fn189 AS Att_1540,
dbo.tnS_S0001.fn190 AS Att_1560,
dbo.tnS_S0001.fn191 AS Att_1570,
dbo.tnS_S0001.fn192 AS Att_1580,
dbo.tnS_S0001.fn193 AS Att_1590,
dbo.tnS_S0001.fn194 AS Att_1600,
dbo.tnS_S0001.fn195 AS Att_1620,
dbo.tnS_S0001.fn196 AS Att_1625,
dbo.tnS_S0001.fn197 AS Att_1650,
dbo.tnS_S0001.fn201 AS Lf_NY,
dbo.tnS_S0001.fn202 AS Twist,
dbo.tnS_S0001.fn203 AS Twist_Direct,
dbo.tnS_S0001.fn204 AS Curl,
dbo.tnS_S0001.fn205 AS TowerNo,
dbo.tnS_S0001.fn188 AS CCUTOFF,
dbo.tnS_S0001.fn096 AS M10_1550,
dbo.tnS_S0001.fn098 AS M10_1625,
dbo.tnS_S0001.fn092 AS M30_1550,
dbo.tnS_S0001.fn093 AS M30_1625,
Left([F_FibreNo],9) AS drum,
Left([F_FibreNo],10) AS f10,
Left([F_FibreNo],8) AS rod,
dbo.tnS_S0001.fn094 AS M20_1550,
dbo.tnS_S0001.fn097 AS M20_1625,
dbo.tnS_S0001.fn122 AS AttMISL1310,
dbo.tnS_S0001.fn123 AS AttMISL1383,
dbo.tnS_S0001.fn124 AS AttMISL1550,
dbo.tnS_S0001.Location,
dbo.tnS_S0001.fn126 AS [minAtt1310-1383],
dbo.tnS_S0001.fn125 AS week,
dbo.tnS_S0001.fn120 AS [盘具],
dbo.tnS_S0001.fn209 AS Status,
Left([F_FibreNo],10) AS fn10,
dbo.tnS_S0001.fn215 AS D2_bn,
dbo.tnS_S0001.fn036 AS Preformtype,
dbo.tnS_S0001.fn212,
dbo.tnS_S0001.fn218

FROM
dbo.tnS_S0001
WHERE
dbo.tnS_S0001.F_ProdTime > '201221000000'
ORDER BY
dbo.tnS_S0001.F_ProdTime DESC



"""


# bookname = r"d:\forAllinOne\临时入库数据.xlsx"

crsr.execute(sql5)   #执行sql1语句
# row = crsr.fetchone()  #读取查询结果,
alldata = crsr.fetchall()
# print(row)
print("========SUCCESS to  get  all data FROM sqlserver========")


# 
# print("====================查询结果写入sqlite")
import sqlite3
cn = sqlite3.connect(r'G:/98sqlite/tns0001.db3')
    # 设定表名称及结构
sql6 = '''
    CREATE TABLE IF NOT EXISTS TNS0001
        (PRODUCTNO TEXT , 
        F_DeptNo TEXT,
        F_BatchNo TEXT,x
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
        fn212 TEXT
       
cn.execute(sql6)




print("===完成create table in sqlite===========")



for t in alldata:
    cn.execute("insert into TNS0001 values (?,?,?,?,?,?,?,?,?,?,?,?,?)", t)
cn.commit()  
cn.close()  
cursor.close
connection.close
print("===完成数据插入e===========")













 