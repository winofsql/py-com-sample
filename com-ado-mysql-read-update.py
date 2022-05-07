# python -m pip install pywin32
import win32com.client

adOpenDynamic = 2
adLockOptimistic = 3

cn = win32com.client.Dispatch("ADODB.Connection")
rs = win32com.client.Dispatch("ADODB.Recordset")

driver = "{MySQL ODBC 8.0 Unicode Driver}"
server = "localhost"
DB = "lightbox"
User = "root"
Pass = ""

ConnectionString = "Provider=MSDASQL;Driver={0};Server={1};DATABASE={2};UID={3};PWD={4};"
ConnectionString = ConnectionString.format(driver,server,DB,User,Pass)
print(ConnectionString)

"""""""""""""""""""""""""""
接続
"""""""""""""""""""""""""""
try:
    cn.Open( ConnectionString )
except Exception as e:
    print( e )
    exit()

rs.CursorType = adOpenDynamic
rs.LockType = adLockOptimistic

rs.Open( "select * from 社員マスタ where 社員コード <= '0004' ", cn )

while not rs.EOF:

    print(rs.Fields("社員コード").Value,end=",")
    print(rs.Fields("氏名").Value, end=",")
    print(rs.Fields("フリガナ").Value, end=",")
    print(rs.Fields("所属").Value, end=",")
    print(rs.Fields("性別").Value, end=",")
    print(rs.Fields("給与").Value, end=",")
    print(rs.Fields("手当").Value, end=",")		# NULL の場合は　None と出力
    print(rs.Fields("管理者").Value, end=",")
    print("{0:%Y/%m/%d}".format(rs.Fields("作成日").Value), end=",")
    print("{0:%Y/%m/%d}".format(rs.Fields("更新日").Value), end=",")
    print("{0:%Y/%m/%d}".format(rs.Fields("生年月日").Value))

    rs.Fields("管理者").Value = "0002"
    rs.Update()

    rs.MoveNext()

"""""""""""""""""""""""""""
接続解除
"""""""""""""""""""""""""""
if cn.State >= 1:
    cn.Close()

print("終了しました")