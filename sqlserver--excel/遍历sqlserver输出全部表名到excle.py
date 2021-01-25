import pymssql #引入pymssql模块
import xlwt

##==============写入第一个数据库10.42.90.92========================
connect = pymssql.connect('10.42.90.92', 'fis', 'fis', 'cab') #服务器名,账户,密码,数据库名

print("连接10.42.90.92成功!")
crsr = connect.cursor() 
# select name from sysobjects where xtype='u'
# select * from sys.tables


#查询全部表名称
# cursor = connect.cursor()   #创建一个游标对象,python里的sql语句都要通过cursor来执行
sql = "select * from sys.tables"
crsr.execute(sql)   #执行sql语句
row = crsr.fetchone()  #读取查询结果,
alldata = crsr.fetchall()
while row:              #循环读取所有结果
##    print("Name=%s, Sex=%s" % (row[0],row[1]))   #输出结果
    row = crsr.fetchone()

    
    
# 写入excel
book = xlwt.Workbook()
sheet1 = book.add_sheet('10.42.90.92')
fields = [field[0] for field in crsr.description]  # 获取所有字段名    
print(fields)
for col,field in enumerate(fields):
    print(col,field)
    sheet1.write(0,col,field)
print ("========完成写表头=========")    
row = 1
for data in alldata:
    #print(data)
##    print ("d%",row) 
    for col,field in enumerate(data):
        sheet1.write(row,col,field)
        #print(type(row))
        #print（row）
        #print（col）
        #print（field）
    row += 1
crsr.close()    
print("======完成=10.42.90.92========")


##==============写入第二个数据库10.42.90.92========================
connect = pymssql.connect('10.42.90.92', 'fis', 'fis', 'cab') #服务器名,账户,密码,数据库名

print("连接成功!")
crsr = connect.cursor() 
# select name from sysobjects where xtype='u'
# select * from sys.tables


#查询全部表名称
# cursor = connect.cursor()   #创建一个游标对象,python里的sql语句都要通过cursor来执行
sql = "select * from sys.tables"
crsr.execute(sql)   #执行sql语句
row = crsr.fetchone()  #读取查询结果,
alldata = crsr.fetchall()
while row:              #循环读取所有结果
##    print("Name=%s, Sex=%s" % (row[0],row[1]))   #输出结果
    row = crsr.fetchone()

    
    
# 写入excel
book = xlwt.Workbook()
sheet1 = book.add_sheet('10.42.90.92')
fields = [field[0] for field in crsr.description]  # 获取所有字段名    
print(fields)
for col,field in enumerate(fields):
    print(col,field)
    sheet1.write(0,col,field)
print ("========完成写表头=========")    
row = 1
for data in alldata:
    #print(data)
##    print ("d%",row) 
    for col,field in enumerate(data):
        sheet1.write(row,col,field)
        #print(type(row))
        #print（row）
        #print（col）
        #print（field）
    row += 1
crsr.close()    
print("======完成=10.42.90.92========")




    
book.save("database_sqlseve_table_list.xls")
print("========完成写入xls=========")



connect.close()
