' �z��ɔz�������
' 2016/9/16 y.ikeda
Option Explicit

Dim ary(2)
ary(0) = 1
ary(1) = Array("a","b")
ary(2) = 2

msgbox ary(1)(1)
' b
' �ł����I�I�ł���Ƃ͎v��Ȃ������E�E�E

msgbox UBound(ary(1))


' �Ē�`�͂ł��Ȃ��͗l�E�E�E
'Redim Preserve ary(1)(3)

