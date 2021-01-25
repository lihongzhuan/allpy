# -*- coding: utf-8 -*-
'''
Created on Tue Dec 22 14:56:12 2020

@author: 0100904
'''
# /===============STEP1：/建立sqlite数据库======
import sqlite3

def createDataBase(dbname):
    cn = sqlite3.connect(dbname)
    cn.execute('''CREATE TABLE IF NOT EXISTS DRUMS
    (
    PRODUCTNO TEXT PRIMARY KEY,
    CUTFROM TEXT,
    LENGTH INT,
    MACHINE TEXT,
    STATUS TEXT,
    DRAWING_END DATE,
    DRAWING_BEGIN DATE,
    subtype TEXT,
    OUTPUTTYPE DATE,
    TENSION INT,
    平均拉丝速度 INT,
    recoverytime INT,
    note TEXT,
    weekno TEXT,
    DRUMID TEXT,
    datefisish DATE,
    blacklist TEXT,
    CLADDIASUB TEXT,
    CLADDIA FLOAT,
    xxxx TEXT,
    jump TEXT,
    towerexpansion TEXT,
    PCREASON TEXT,
    SPEEDDOWNREASON TEXT,
    pistonchange TEXT,
    pistonchange_reason TEXT,
    DRUMSN TEXT,
    PISTONSN TEXT ,
    棒子类型 TEXT);''')
    print("===完成table1===")
    
    
    
if __name__ == '__main__':
    createDataBase('DRUM制造信息.db3')
















