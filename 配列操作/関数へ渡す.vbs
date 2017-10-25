' 配列を関数へ渡す

Option Explicit


' 配列定義
Dim array1
array1 = Array("a","b","c")		'Array関数を使用

' 呼ぶ
test(array1)


' 関数定義
Sub test(pArray)
	' 数を数えてみる
	msgbox UBound(pArray), vbOkOnly, "UBound(array)"
	' → 2
	
	' 出力してみる
	msgbox pArray(0) & "-" & pArray(1) & "-" & pArray(2), vbOkOnly, "elements"
	' → a-b-c
End Sub

