Option Explicit


Dim str1
str1 = "test1"

Dim str3
str3 = "test3-1"

' Sub呼び出し
Call sub_B("呼べるかな")

'MsgBox str2
' ↑エラーになる

MsgBox str3

' Sub呼び出し
Call sub_C()



Sub sub_B(param1)
	MsgBox(param1)
	MsgBox(str1)
	
	Dim str2
	str2 = "test2"
	
	str3 = "test3-2"
	
End Sub

Sub sub_C()
	MsgBox(str3)
	
End Sub


' 結論
' 関数外で定義したものは、関数内で読み取り/変更可能
' 関数内で定義したものは、関数外で実行時エラーになる。Option Explicitでも。


