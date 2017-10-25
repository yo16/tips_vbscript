Option Explicit

Dim myArray(2)
myArray(0) = "い"
myArray(1) = "ろ"
myArray(2) = "は"

'2つ目の引数""は、区切り文字なしって意味。
'引数なしの場合は、スペースで区切られる。
MsgBox Join(myArray,"")
