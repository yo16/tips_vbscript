Option Explicit

' for continue�̂悤�Ȃ��̂�
' VBS�ɂ͂Ȃ�


Dim idx
Dim tmpNumber
tmpNumber = 0

For idx = 1 to 10 : Do
	If (idx > 5) Then Exit Do
	tmpNumber = tmpNumber + 1
Loop Until 1: Next

MsgBox "tmpNumber is "&tmpNumber


' ���L��For��Do���P�s�ŏ���������
'	For idx = 1 to 10
'		Do
'			If (idx > 5) Then Exit Do
'			tmpNumber = tmpNumber + 1
'		Loop Until 1
'	Next


