import pyodbc
import faker




# 准备模拟数据
fake = faker.Faker('zh_CN')
# 设置种子值，不设的话每次随机的都不一样

##fake.Faker.seed(47)

db_file_location = r'D:\users\基础链接表.mdb'
# 这里用的是Python3.5的语法，如果是低版本Python的话需要改成普通方式
connection = pyodbc.connect(rf'Driver={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={db_file_location};')

connection.autocommit = True
cursor = connection.cursor()
print("====================完成数据库" + db_file_location + "连接======================")
# 第一次创建表，将其设置为False

print("====================开始创建表数据库连接======================")
# 第一次创建表，将其设置为False

##create_table_sql = '''\
##create table user8
##(
##  id        autoincrement primary key,
##  username  varchar(255) unique,
##  nickname  varchar(255) not null,
##  password  varchar(20)  not null,
##  address   varchar(255),
##  birthday  date,
##  company   varchar(30),
##  job       varchar(20),
##  telephone varchar(14)
##)
##'''1000_RIC组装表

##sql1 ='''\
##DELETE [0000_REJCT汇总-1].筛选长度之合计
##FROM [0000_REJCT汇总-1]
##WHERE ((([0000_REJCT汇总-1].筛选长度之合计) Like "*"))
##'''
##cursor.execute(sql1)




sqltemp ='''\
select top 10 *
from
0000RIC组装表


'''




for row in cursor.execute(sqltemp):
  print(row)

print("====================完成sql 1 ======================")



##select_public_servant_sql = '''\
##INSERT INTO 1000_RIC组装表 ( CYLINDER, SCRNUM, SCRNO, RICNUM, ITEM, RIC_END, WASH_END_WEIGHT, COREROD, 表达式2, deltah1, deltah2, SCRETCHOD, LINK, SUBTYPE, DRAWLENGTH, REMARK, 表达式1, NEXTOUTDATE, addtime )
##SELECT AAA_FISRICASSEMBLEPARA.CYLINDER, AAA_FISRICASSEMBLEPARA.SCRNUM, AAA_FISRICASSEMBLEPARA.SCRNO, AAA_FISRICASSEMBLEPARA.RICNUM, AAA_FISRICASSEMBLEPARA.ITEM, AAA_FISRICASSEMBLEPARA.RIC_END, AAA_TUBE.WASH_END_WEIGHT, AAA_FISSCRPARA.COREROD, Round([CYLINDERLENGTH],0) AS 表达式2, AAA_FISRICASSEMBLEPARA.DH AS deltah1, AAA_FISRICASSEMBLEPARA.CONELENGTH AS deltah2, AAA_FISRICASSEMBLEPARA.SCRETCHOD, Left([CYLINDER],8) AS LINK, AAA_FISRICASSEMBLEPARA.SUBTYPE, AAA_FISRICASSEMBLEPARA.DRAWLENGTH, AAA_FISRICASSEMBLEPARA.REMARK, Left([CYLINDER],8) AS 表达式1, AAA_FISRICASSEMBLEPARA.NEXTOUTDATE, Now() AS 表达式3
##FROM AAA_FISSCRPARA RIGHT JOIN (AAA_FISRICASSEMBLEPARA LEFT JOIN AAA_TUBE ON AAA_FISRICASSEMBLEPARA.CYLINDER = AAA_TUBE.TUBE_ID) ON AAA_FISSCRPARA.SCRNUM = AAA_FISRICASSEMBLEPARA.SCRNO
##WHERE (((AAA_FISRICASSEMBLEPARA.SCRNUM) Like "*") AND ((AAA_FISRICASSEMBLEPARA.RIC_END)>#6/1/2020#));




##for row in cursor.execute(sql1):
##  print(row)



##table_exists = False
##if not table_exists:
##    with connection.cursor() as cursor:
##        cursor.execute(create_table_sql)
##        
### 添加数据
##with connection.cursor() as cursor:
##    for _ in range(3000):
##        cursor.execute(insert_table_sql, (fake.pystr(min_chars=6, max_chars=10),
##                                          fake.name(),
##                                          fake.password(length=10),
##                                          fake.address(),
##                                          fake.date_of_birth(minimum_age=0, maximum_age=120),
##                                          fake.company(),
##                                          fake.job(),
##                                          fake.phone_number()))
##
##    # 查询一下所有公务员
##    cursor.execute(select_public_servant_sql)
##    results = cursor.fetchall()
##    for row in results:
##        print(row[0], row[1], row[2], row[3], row[4], row[5], row[6], row[7], sep='\t')
cursor.close
connection.close

print("====================关闭数据库，成功完成全部工作======================")

