Option Explicit

Dim arynum
arynum = 3


' ���̏����������ƁA�v�f�P��'3'�Ƃ����l�����邾��
Dim ary1
ary1 = Array(arynum)

msgbox UBound(ary1), vbOkOnly, "ary1"
' �� 0

Dim i
For i=0 to UBound(ary1)
	msgbox i&":"&ary1(i), vbOkOnly, "ary1"
Next



' ���̏����������ƁAUBound��3�ɂȂ��āA�l�͑S����
' �ł��ϐ��͎g���Ȃ�.
'Dim ary2(arynum)	' �R���p�C���G���[�ɂȂ�
Dim ary2(3)
msgbox UBound(ary2), vbOkOnly, "ary2"
' �� 3

For i=0 to UBound(ary2)
	msgbox i&":"&ary2(i), vbOkOnly, "ary2"
Next



' ReDim���g�����ƂőS������
Dim ary3
ary3 = Array()
msgbox UBound(ary3), vbOkOnly, "ary3-1"
ReDim ary3(arynum)		' �����Ŏw�肷��̂�UBound�l
msgbox UBound(ary3), vbOkOnly, "ary3-2"
' �� 3�E�E�E3�v�f�ł͂Ȃ��A4�v�f�ɂȂ邱�Ƃɒ��ӁI

For i=0 to UBound(ary3)
	msgbox i&":"&ary3(i), vbOkOnly, "ary3"
Next
