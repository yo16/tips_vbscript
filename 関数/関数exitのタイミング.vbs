Option Explicit

' Function�Ăяo��

Dim returnValue
returnValue = func_A(1, 2)




Function func_A(paramA, paramB)
	' �߂�l�ݒ�
	func_A = paramA+paramB
	
	
	Exit Function		' �����ꂪ�K�v
	
	msgbox "�R�R�ʂ�H"
	
End Function
