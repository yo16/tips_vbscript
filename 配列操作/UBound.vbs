Option Explicit

Dim array1
array1 = Array("a","b","c")		'Array関数を使用

' 最後のインデックスを返す
MsgBox UBound( array1 )



' ゼロの時は-1
Dim array2
array2 = Array()
MsgBox UBound( array2 )
