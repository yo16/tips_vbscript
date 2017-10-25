Option Explicit

Dim dbCon
Dim dbRec

Set dbCon = CreateObject("ADODB.Connection")
'dbCon.Open "Provider=Microsoft.ACE.OLEDB.12.0;Jet OLEDB:Engine Type=6;Data Source=test2.accdb;"

'Microsoft Access データベース エンジン 2010 再頒布可能コンポーネント
' を導入して、ダブルクリックで実行できるようになった。
'↓
dbCon.Open "Provider=Microsoft.Ace.OLEDB.12.0;Data Source=test2.accdb;"

'dbCon.Open "Provider=Microsoft.Ace.OLEDB.12.0;Data Source=test2.accdb;Persist Security Info=False;"
'dbCon.Open "driver={Microsoft Office 12.0 Access Database Engine OLE EB Provider};Data Source=test2.accdb;"
'dbCon.Open "driver={Microsoft Access Driver (*.mdb)};Data Source=test2.accdb;"
'dbCon.Open "driver={Microsoft Access Driver (*.mdb, *.accdb)};Data Source=test2.accdb;"
'dbCon.Open "driver={Microsoft Access Driver (*.accdb)};DBQ=test2.accdb;"
'dbCon.Open "driver={Microsoft Access Driver (*.mdb)};Dbq=test2.accdb;"

' 再配布可能コンポーネントを入れても、わけのわからないエラーが出て接続できない。
' ↓
'dbCon.Open "driver={Microsoft Access Driver (*.mdb, *.accdb)};Dbq=test2.accdb;"


'dbCon.Open "driver={Microsoft Access Driver (*.accdb)};Dbq=test2.accdb;"


' SQL実行
Set dbRec = dbCon.Execute("select * from table1")

Do Until dbRec.EOF
	WScript.Echo dbRec("ID")
	
	dbRec.MoveNext
Loop

dbRec.Close

dbCon.Close

msgbox "完了!"
