
Dim str

str = "123"
If IsNumeric(str) Then
	MsgBox str & "�͐���"	' ���s���ʁF����
Else
	MsgBox str & "�͐����łȂ�"
End If


str = "-123"
If IsNumeric(str) Then
	MsgBox str & "�͐���"	' ���s���ʁF����
Else
	MsgBox str & "�͐����łȂ�"
End If


str = "-1.23"
If IsNumeric(str) Then
	MsgBox str & "�͐���"	' ���s���ʁF����
Else
	MsgBox str & "�͐����łȂ�"
End If


str = "-1.2.3"
If IsNumeric(str) Then
	MsgBox str & "�͐���"
Else
	MsgBox str & "�͐����łȂ�"	' ���s���ʁF�����łȂ�
End If
