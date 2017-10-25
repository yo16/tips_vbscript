' 多次元配列を拡張

Option Explicit

' 後でReDimしたい場合は、この書き方（またはary1 = Array()）
Dim ary1
ReDim ary1(1,2)
ary1(0,0) = "0-0"
ary1(0,1) = "0-1"
ary1(0,2) = "0-2"
ary1(1,0) = "1-0"
ary1(1,1) = "1-1"
ary1(1,2) = "1-2"

' 最後の次元しか、拡張できない（仕様・制限）
ReDim Preserve ary1(1,3)
ary1(0,3) = "0-3"
ary1(1,3) = "1-3"

' 確認
Dim i, j
For i=0 to UBound(ary1,1)	' ary1の1次元目のUBound
For j=0 to UBound(ary1,2)	' ary1の2次元目のUBound
	MsgBox ary1(i,j)
	' → "0-0"、"0-1"、"0-2"、・・・"1-2"、"1-3"
Next
Next

