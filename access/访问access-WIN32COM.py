"""
通过pyWin32连接Access数据库
实测已通过
"""
import win32com.client
import sys

conn = win32com.client.Dispatch(r'ADODB.Connection')

# 第一种连接串
# DSN = r"PROVIDER=Microsoft.Jet.OLEDB.4.0;DATA SOURCE=d:\Python code\员工信息.mdb;Persist Security Info=False;Jet " \
#      r"OLEDB:Database Password=123456"
# 第二种连接串
# 需要安装AccessDatabaseEngine
# 64位系统下载地址：(https://download.microsoft.com/download/E/4/2/E4220252-5FAE-4F0A-B1B9-0B48B5FBCCF9/AccessDatabaseEngine_X64.exe)
# 32位系统下载地址：(https://download.microsoft.com/download/E/4/2/E4220252-5FAE-4F0A-B1B9-0B48B5FBCCF9/AccessDatabaseEngine.exe)

# DSN = r"Provider=Microsoft.ACE.OLEDB.12.0;DATA SOURCE=c:\t.mdb;Persist Security Info=False;"

DSN =r'PROVIDER=Microsoft.ACE.OLEDB.12.0;DATA SOURCE=c:\t.mdb;'   
conn.Open(DSN)
print("connect...")

rs = win32com.client.Dispatch(r'ADODB.Recordset')

rs.Open('Select * FROM tbLogin', conn, 1, 3)
if rs.recordcount == 0:
    sys.exit()

rs.MoveFirst()
print(rs.recordcount)

while not rs.EOF:
    print(rs.Fields.Item('ID').Value)
    print("-------------------------------------")
    rs.MoveNext()

print("Record Count: ", rs.recordcount)

rs.Close()
