Option Explicit


Dim str1
str1 = "test1"

Dim str3
str3 = "test3-1"

' Sub�Ăяo��
Call sub_B("�Ăׂ邩��")

'MsgBox str2
' ���G���[�ɂȂ�

MsgBox str3

' Sub�Ăяo��
Call sub_C()



Sub sub_B(param1)
	MsgBox(param1)
	MsgBox(str1)
	
	Dim str2
	str2 = "test2"
	
	str3 = "test3-2"
	
End Sub

Sub sub_C()
	MsgBox(str3)
	
End Sub


' ���_
' �֐��O�Œ�`�������̂́A�֐����œǂݎ��/�ύX�\
' �֐����Œ�`�������̂́A�֐��O�Ŏ��s���G���[�ɂȂ�BOption Explicit�ł��B


