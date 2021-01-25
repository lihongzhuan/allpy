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

print("====================开始处理数据库数据======================")
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


### 修改数据库中int类型的值
##value = 10
##SQL = "UPDATE goods " \
##   "SET lowestStock=" + str(value) + " " \
##   "WHERE goodsId='0005'"
## 
### 删除表users
##crsr.execute("DROP TABLE users")
### 创建新表 users
##crsr.execute('CREATE TABLE users (login VARCHAR(8),userid INT, projid INT)')
### 给表中插入新数据
##crsr.execute("INSERT INTO users VALUES('Linda',211,151)")
## 
##''''''
### 更新数据
##crsr.execute("UPDATE users SET projid=1 WHERE userid=211")
## 
### 删除行数据
##crsr.execute("DELETE FROM goods WHERE goodNum='0001'")
## 
### 打印查询的结果
##for row in crsr.execute("SELECT * from users"):
##  print(row)
##



######create table tab1 as select * from tab2


##sqltemp ='''\
##select top 10 *
##from
##0000RIC组装表
##
##
##'''
##
##
##
##
##for row in cursor.execute(sqltemp):
##  print(row)
##
##print("====================完成sql temp ======================")
##
##
##
##sql1 ='''\
##DELETE FROM 1000REJCT汇总
##'''
##cursor.execute(sql1)
##
##
##
##
##
##print("====================完成sql 1 ======================")




##cursor.execute("DELETE FROM 0002sct汇总")



##sql2 ='''\
##CREAT TABLE 0002sct AS
##(SELECT AAA_FIBRE_SCT_NEW.DATETIME, AAA_FIBRE_SCT_NEW.DRUM AS SCT的drum号, Left([DRUM],9) AS 拉丝Drum号, Left([DRUM],8) AS ROD, AAA_FIBRE_SCT_NEW.DRAWTOWER, AAA_FIBRE_SCT_NEW.CODE,
##CDate("20" & Left([DATETIME],2) & "-" & Mid([DATETIME],3,2) & "-" & Mid([DATETIME],5,2)) AS sct_date,
##DatePart("ww",[sct_date]+7,6,2)+2000-1 AS WEEKNUM, AAA_FIBRE_SCT_NEW.LENGTH1, DatePart("m",[sct_date]) AS [month],
##Year([sct_date]+7)-2000 AS [year], AAA_FIBRE_SCT_NEW.OUTPUTTYPE, AAA_FIBRE_SCT_NEW.CUTFROM, AAA_FIBRE_SCT_NEW.SCTEND,
##IIf(Len([DRUM])=9,IIf(Len([FIBRE])=13,[SETSPEED],[LENGTH]),0) AS 筛选长度, IIf([CODE]<>"000",[LENGTH],0) AS REJECT_LEN,
##IIf([CODE]="420",[LENGTH],0) AS 420_LEN, IIf([CODE]="421",[LENGTH],0) AS 421_LEN, IIf([CODE]="422",[LENGTH],0) AS 422_LEN, IIf([CODE]="425",[LENGTH],0) AS 425_LEN, IIf([CODE]="427",
##[LENGTH],0) AS 427_LEN, IIf([CODE]="428",[LENGTH],0) AS 428_LEN, IIf([CODE]="450",[LENGTH],0) AS 450_LEN, IIf([CODE]="452",[LENGTH],0) AS 452_LEN, IIf([CODE]="453",[LENGTH],0) AS 453_LEN,
##IIf([CODE]="454",[LENGTH],0) AS 454_LEN, IIf([CODE]="410",[LENGTH],0) AS 410_LEN, IIf([CODE]="411",[LENGTH],0) AS 411_LEN, IIf([CODE]="413",[LENGTH],0) AS 413_LEN, IIf([CODE]="414",[LENGTH],0) AS 414_LEN,
##IIf([CODE]="415",[LENGTH],0) AS 415_LEN, IIf([CODE]="416",[LENGTH],0) AS 416_LEN, IIf([CODE]="417",[LENGTH],0) AS 417_LEN, IIf([CODE]="418",[LENGTH],0) AS 418_LEN, IIf([CODE]="412",[LENGTH],0) AS 412_LEN,
##IIf([CODE]="431",[LENGTH],0) AS 431_LEN, IIf([CODE]="432",[LENGTH],0) AS 432_LEN, IIf([CODE]="433",[LENGTH],0) AS 433_LEN, IIf([CODE]="434",[LENGTH],0) AS 434_LEN, IIf([CODE]="436",[LENGTH],0) AS 436_LEN,
##IIf([CODE]="437",[LENGTH],0) AS 437_LEN, IIf([CODE]="438",[LENGTH],0) AS 438_LEN, IIf([CODE]="354",[LENGTH],0) AS 354_LEN, IIf([CODE]="355",[LENGTH],0) AS 355_LEN, IIf([CODE]="356",[LENGTH],0) AS 356_LEN,
##IIf([CODE]="357",[LENGTH],0) AS 357_LEN, IIf([CODE]="358",[LENGTH],0) AS 358_LEN, IIf([CODE]="439",[LENGTH],0) AS 439_LEN, IIf([CODE]="455",[LENGTH],0) AS 455_LEN, IIf([CODE]="462",[LENGTH],0) AS 462_LEN,
##IIf([CODE]="499",[LENGTH],0) AS 499_LEN, AAA_FIBRE_SCT_NEW.ITEM, AAA_FIBRE_SCT_NEW.SETLENGTH, AAA_FIBRE_SCT_NEW.SETSPEED, AAA_FIBRE_SCT_NEW.LENGTH, AAA_FIBRE_SCT_NEW.FIBRE, AAA_FIBRE_SCT_NEW.SUBTYPE,
##AAA_FIBRE_SCT_NEW.SCTER, AAA_FIBRE_SCT_NEW.OPER 
##FROM AAA_FIBRE_SCT_NEW
##WHERE (((AAA_FIBRE_SCT_NEW.DRUM) Not Like "??R*" And (AAA_FIBRE_SCT_NEW.DRUM) Like "*") AND ((CDate("20" & Left([DATETIME],2) & "-" & Mid([DATETIME],3,2) & "-" & Mid([DATETIME],5,2)))>Now()-3))
##ORDER BY CDate("20" & Left([DATETIME],2) & "-" & Mid([DATETIME],3,2) & "-" & Mid([DATETIME],5,2)))
##
##
##'''
##
####cursor.execute(sql2)
##




##print("====================完成sql2 ======================")

sql3 ='''\
INSERT INTO 0000RIC组装表(CYLINDER, SCRNUM, SCRNO, RICNUM, ITEM, RIC_END, 表达式2, deltah1, deltah2, SCRETCHOD, LINK, SUBTYPE, DRAWLENGTH, REMARK, 表达式1, NEXTOUTDATE, addtime )
SELECT AAA_FISRICASSEMBLEPARA.CYLINDER, AAA_FISRICASSEMBLEPARA.SCRNUM, AAA_FISRICASSEMBLEPARA.SCRNO, AAA_FISRICASSEMBLEPARA.RICNUM, AAA_FISRICASSEMBLEPARA.ITEM, AAA_FISRICASSEMBLEPARA.RIC_END,
Round([CYLINDERLENGTH],0) AS 表达式2, AAA_FISRICASSEMBLEPARA.DH AS deltah1, AAA_FISRICASSEMBLEPARA.CONELENGTH AS deltah2, AAA_FISRICASSEMBLEPARA.SCRETCHOD, Left([CYLINDER],8) AS LINK, AAA_FISRICASSEMBLEPARA.SUBTYPE,
AAA_FISRICASSEMBLEPARA.DRAWLENGTH, AAA_FISRICASSEMBLEPARA.REMARK, Left([CYLINDER],8) AS 表达式1, AAA_FISRICASSEMBLEPARA.NEXTOUTDATE, Now() AS 表达式3
FROM AAA_FISRICASSEMBLEPARA
WHERE AAA_FISRICASSEMBLEPARA.RIC_END>#6/1/2020#



'''




cursor.execute(sql3)

print("====================完成sql3 ======================")










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

