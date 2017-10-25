
Dim str

str = "123"
If IsNumeric(str) Then
	MsgBox str & "は数字"	' 実行結果：数字
Else
	MsgBox str & "は数字でない"
End If


str = "-123"
If IsNumeric(str) Then
	MsgBox str & "は数字"	' 実行結果：数字
Else
	MsgBox str & "は数字でない"
End If


str = "-1.23"
If IsNumeric(str) Then
	MsgBox str & "は数字"	' 実行結果：数字
Else
	MsgBox str & "は数字でない"
End If


str = "-1.2.3"
If IsNumeric(str) Then
	MsgBox str & "は数字"
Else
	MsgBox str & "は数字でない"	' 実行結果：数字でない
End If
