Option Explicit

' Function�Ăяo��

Dim returnValue
returnValue = func_A(1, 2)
msgbox "Function�Ăׂ�����"&returnValue


' Sub�Ăяo��
Call sub_B("�Ăׂ邩��")




Function func_A(paramA, paramB)
	' �߂�l�ݒ�
	func_A = paramA+paramB
End Function

Sub sub_B(param1)
	MsgBox(param1)
End Sub

