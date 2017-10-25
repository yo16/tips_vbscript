Option Explicit
' 配列要素を再定義して要素を増やす
' 2015/7/30

' 配列を定義
' 　　　後で再定義したい場合は、Array()で作る必要がある
Dim ary
ary = Array("a", "b", "c")
msgbox toStr(ary)
' a-b-c-

' 再定義
' 　　　最後の要素IDを指定する
' 　　　元の要素を保持したいときはPreserveをつける
ReDim Preserve ary(4)
ary(3) = "d"
ary(4) = "e"
'ary(5) = "f"	' ・・・実行時エラーになる
msgbox toStr(ary)
' a-b-c-d-e-


Function toStr(ar)
	Dim str
	str = ""
	Dim i
	For i=0 to UBound(ar)
		str = str & ar(i) & "-"
	Next
	
	toStr = str
End Function
