' �������z����g��

Option Explicit

' ���ReDim�������ꍇ�́A���̏������i�܂���ary1 = Array()�j
Dim ary1
ReDim ary1(1,2)
ary1(0,0) = "0-0"
ary1(0,1) = "0-1"
ary1(0,2) = "0-2"
ary1(1,0) = "1-0"
ary1(1,1) = "1-1"
ary1(1,2) = "1-2"

' �Ō�̎��������A�g���ł��Ȃ��i�d�l�E�����j
ReDim Preserve ary1(1,3)
ary1(0,3) = "0-3"
ary1(1,3) = "1-3"

' �m�F
Dim i, j
For i=0 to UBound(ary1,1)	' ary1��1�����ڂ�UBound
For j=0 to UBound(ary1,2)	' ary1��2�����ڂ�UBound
	MsgBox ary1(i,j)
	' �� "0-0"�A"0-1"�A"0-2"�A�E�E�E"1-2"�A"1-3"
Next
Next

