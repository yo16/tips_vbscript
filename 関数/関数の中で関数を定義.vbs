Option Explicit

Call sub_A()

msgbox "���Ғʂ�ɂ͂ł��Ȃ������B",,"�c�O���O"

Sub sub_A()
	Dim num1,num2
	
	num1 = 1
	num2 = 2

	msgbox func_add(num1,num2)

'�����ɂ͒�`�ł��Ȃ������B2001/11/28
'	Function func_add(p1,p2)
'		func_add = p1+p2
'	End Function

End Sub

' �O�ɏo������OK���o���B
	Function func_add(p1,p2)
		func_add = p1+p2
	End Function
