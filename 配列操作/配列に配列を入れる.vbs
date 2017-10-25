' 配列に配列を入れる
' 2016/9/16 y.ikeda
Option Explicit

Dim ary(2)
ary(0) = 1
ary(1) = Array("a","b")
ary(2) = 2

msgbox ary(1)(1)
' b
' できた！！できるとは思わなかった・・・

msgbox UBound(ary(1))


' 再定義はできない模様・・・
'Redim Preserve ary(1)(3)

