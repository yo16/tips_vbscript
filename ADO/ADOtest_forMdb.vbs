Option Explicit

Dim dbCon
Dim dbRec

Set dbCon = CreateObject("ADODB.Connection")
dbCon.Open "driver={Microsoft Access Driver (*.mdb)};DBQ=test.mdb;"

' SQLé¿çs
Set dbRec = dbCon.Execute("select * from table1")

Do Until dbRec.EOF
	WScript.Echo dbRec("ID")
	
	dbRec.MoveNext
Loop

dbRec.Close

dbCon.Close

msgbox "äÆóπ!"
