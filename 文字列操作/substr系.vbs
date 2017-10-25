'*************************************
'substr系の関数
'
' Left、Mid、Rightがある
'*************************************

Option Explicit

Dim testStr
testStr = "abcdefg"

MsgBox "Left:" & Left(testStr, 3)
' abc
MsgBox "Mid:" & Mid(testStr, 2, 3)
' bcd
MsgBox "Right:" & Right(testStr, 3)		' 開始位置から右
' efg



'文字列の、頭１桁を取る
MsgBox "頭１桁を取る：" & Right(testStr, Len(testStr)-1)
' bcdefg
