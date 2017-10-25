Option Explicit

' 多次元配列の定義方法
Dim ary1(1,2)
ary1(0,0) = "0-0"
ary1(0,1) = "0-1"
ary1(0,2) = "0-2"
ary1(1,0) = "1-0"
ary1(1,1) = "1-1"
ary1(1,2) = "1-2"

Dim i, j
For i=0 to UBound(ary1,1)	' ary1の1次元目のUBound
For j=0 to UBound(ary1,2)	' ary1の2次元目のUBound
	MsgBox ary1(i,j)
	' → "0-0"、"0-1"、"0-2"、"1-0"、"1-1"、"1-2"
Next
Next



' メモ
Dim d1, d2
d1 = 1
d2 = 2
'Dim ary2(d1,d2)
' 変数を使った定義はできない
