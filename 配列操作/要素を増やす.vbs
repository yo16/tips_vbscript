Option Explicit


Dim array1
array1 = Array("a","b","c")
msgbox array1(2)
' この後ろに要素をもう１個付け加えたいなぁ。。


' ダメ
'array1(3) = "d"
'msgbox array1(3)

' ダメ
'array1 = Array(array1(0), array1(1), array1(2), "d")
'msgbox array1(3)

' ＯＫ!!!
Dim arrayTmp
arrayTmp = array1
array1 = null	'********ここが重要*******
array1 = Array(arrayTmp(0), arrayTmp(1), arrayTmp(2), "d")
msgbox array1(3)


