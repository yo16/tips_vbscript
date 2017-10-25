' 関数に渡してみるテスト


' 簡単にできちゃった 2007/04/24 



Option Explicit


' 配列定義
Dim array1
array1 = Array("a","b","c")		'Array関数を使用

' 呼ぶ
test(array1)


' 関数定義
Sub test(pArray)
	' 数を数えてみる
	msgbox UBound(pArray)
	
	' 出力してみる
	msgbox pArray(0) & "-" & pArray(1) & "-" & pArray(2)


End Sub

