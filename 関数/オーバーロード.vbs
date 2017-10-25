Option Explicit
' 関数のオーバーロードはできない

AAA "a"		' 実行時エラーになる
AAA "a", "z"


Sub AAA(param1)
	msgbox param1
End Sub

Sub AAA(param1, param2)
	msgbox param1 & "---" & param2
End Sub
