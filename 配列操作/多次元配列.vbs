Option Explicit

' �������z��̒�`���@
Dim ary1(1,2)
ary1(0,0) = "0-0"
ary1(0,1) = "0-1"
ary1(0,2) = "0-2"
ary1(1,0) = "1-0"
ary1(1,1) = "1-1"
ary1(1,2) = "1-2"

Dim i, j
For i=0 to UBound(ary1,1)	' ary1��1�����ڂ�UBound
For j=0 to UBound(ary1,2)	' ary1��2�����ڂ�UBound
	MsgBox ary1(i,j)
	' �� "0-0"�A"0-1"�A"0-2"�A"1-0"�A"1-1"�A"1-2"
Next
Next



' ����
Dim d1, d2
d1 = 1
d2 = 2
'Dim ary2(d1,d2)
' �ϐ����g������`�͂ł��Ȃ�
