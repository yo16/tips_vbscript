Option Explicit

Dim dbCon
Dim dbRec

Set dbCon = CreateObject("ADODB.Connection")
'dbCon.Open "Provider=Microsoft.ACE.OLEDB.12.0;Jet OLEDB:Engine Type=6;Data Source=test2.accdb;"

'Microsoft Access �f�[�^�x�[�X �G���W�� 2010 �ĔЕz�\�R���|�[�l���g
' �𓱓����āA�_�u���N���b�N�Ŏ��s�ł���悤�ɂȂ����B
'��
dbCon.Open "Provider=Microsoft.Ace.OLEDB.12.0;Data Source=test2.accdb;"

'dbCon.Open "Provider=Microsoft.Ace.OLEDB.12.0;Data Source=test2.accdb;Persist Security Info=False;"
'dbCon.Open "driver={Microsoft Office 12.0 Access Database Engine OLE EB Provider};Data Source=test2.accdb;"
'dbCon.Open "driver={Microsoft Access Driver (*.mdb)};Data Source=test2.accdb;"
'dbCon.Open "driver={Microsoft Access Driver (*.mdb, *.accdb)};Data Source=test2.accdb;"
'dbCon.Open "driver={Microsoft Access Driver (*.accdb)};DBQ=test2.accdb;"
'dbCon.Open "driver={Microsoft Access Driver (*.mdb)};Dbq=test2.accdb;"

' �Ĕz�z�\�R���|�[�l���g�����Ă��A�킯�̂킩��Ȃ��G���[���o�Đڑ��ł��Ȃ��B
' ��
'dbCon.Open "driver={Microsoft Access Driver (*.mdb, *.accdb)};Dbq=test2.accdb;"


'dbCon.Open "driver={Microsoft Access Driver (*.accdb)};Dbq=test2.accdb;"


' SQL���s
Set dbRec = dbCon.Execute("select * from table1")

Do Until dbRec.EOF
	WScript.Echo dbRec("ID")
	
	dbRec.MoveNext
Loop

dbRec.Close

dbCon.Close

msgbox "����!"
