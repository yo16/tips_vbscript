' ODBCê⁄ë±Ç≈SPOOL
' 2017/2/21 (c) y.ikeda

Option Explicit

Dim outFilePath
outFilePath = "out.csv"

Const SV = "10.20.30.40"
Const DB = "dbname123"
Const PW = "dbpw123"

Dim query
query = _
	"select support_number, status, answerer_name, first_date, upd_date " & _
	"from sp_support_log " & _
	"where del_flg = 0 and " & _
	"( first_date >= to_date('20161101','YYYYMMDD') and status in (4, 5, 6) ) or " & _
	"( upd_date >= to_date('20170221', 'YYYYMMDD') ) " & _
	"order by support_number asc;"

' ê⁄ë±
Dim con
Set con = WScript.CreateObject("ADODB.Connection")
con.Open "Provider=MSDASQL;Driver=PostgreSQL Unicode(x64);UID=postgres;Port=5432" &_
         ";Server=" & SV & ";Database=" & DB & ";PWD=" & PW

' Open
Dim rs
Set rs = WScript.CreateObject("ADODB.Recordset")
rs.Open query, con
If Err.Number <> 0 Then
	con.Close
	MsgBox Err.Description, vbOkOnly, "Opené∏îs"
	WScript.Quit
End If

' CSV Open
Dim objFs
Set objFs = WScript.CreateObject("Scripting.FileSystemObject")
Dim objOut
Set objOut = objFs.CreateTextFile( outFilePath, true )

' åüçıÅïèoóÕ
Dim buf
Dim i
Do While Not rs.EOF
	buf = ""
	For i=0 to rs.Fields.Count - 1
		If buf <> "" Then
			buf = buf & ","
		End If
		buf = buf & """" & rs.Fields(i).Value & """"
	Next
	objOut.Writeline buf
	rs.MoveNext
Loop

' CSV Close
objOut.Close

' DB Close
rs.Close
con.Close
Set con = Nothing
WScript.Echo "èIóπÅI"
