Option Explicit
On Error Resume Next

'はじめ
If Err Then
	MsgBox "エラー！",,"はじめ"
Else
	MsgBox "正常。",,"はじめ"
End If

'エラーを起こしてみる
Err.Raise(10)
If Err Then
	MsgBox "エラー！",,"エラーを起こしてみる"
Else
	MsgBox "正常。",,"エラーを起こしてみる"
End If

'ほっといてみる
If Err Then
	MsgBox "エラー！",,"ほっといてみる"
Else
	MsgBox "正常。",,"ほっといてみる"
End If

'クリアしてみる
Err.Clear
If Err Then
	MsgBox "エラー！",,"クリアしてみる"
Else
	MsgBox "正常。",,"クリアしてみる"
End If
