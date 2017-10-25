Option Explicit
' 配列をコピー

Dim ary1, ary2
ary1 = Array("a","b")
ary2 = ary1

msgbox "UBound:" & UBound(ary2)
' → UBound:1
msgbox "1:" & ary2(0) & " 2:" & ary2(1)
' → 1:a 2:b
