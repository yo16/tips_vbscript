Option Explicit

Dim arynum
arynum = 3


' この書きかただと、要素１つに'3'という値が入るだけ
Dim ary1
ary1 = Array(arynum)

msgbox UBound(ary1), vbOkOnly, "ary1"
' → 0

Dim i
For i=0 to UBound(ary1)
	msgbox i&":"&ary1(i), vbOkOnly, "ary1"
Next



' この書きかただと、UBoundは3になって、値は全部空
' でも変数は使えない.
'Dim ary2(arynum)	' コンパイルエラーになる
Dim ary2(3)
msgbox UBound(ary2), vbOkOnly, "ary2"
' → 3

For i=0 to UBound(ary2)
	msgbox i&":"&ary2(i), vbOkOnly, "ary2"
Next



' ReDimを使うことで全部解決
Dim ary3
ary3 = Array()
msgbox UBound(ary3), vbOkOnly, "ary3-1"
ReDim ary3(arynum)		' ここで指定するのはUBound値
msgbox UBound(ary3), vbOkOnly, "ary3-2"
' → 3・・・3要素ではなく、4要素になることに注意！

For i=0 to UBound(ary3)
	msgbox i&":"&ary3(i), vbOkOnly, "ary3"
Next
