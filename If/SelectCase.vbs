Option Explicit
' Select-Case
' C言語でいうSwitch-Case。
' ただしC言語のように、１つつかまったら順に次へ進むわけではなく
' Caseの範囲が終わったら最後まで飛ぶ仕様。便利。

Dim str
str = "aa"

Select Case str
Case "aa"
	MsgBox "aaですよ"
Case "bb"
	MsgBox "bbですよ"
Case Else
	MsgBox "elseですよ"
End Select
