Option Explicit

Dim myArray
myArray = Array("A", "B", "C")

MsgBox isExists(myArray, "A")
MsgBox isExists(myArray, "B")
MsgBox isExists(myArray, "C")
MsgBox isExists(myArray, "a")
MsgBox isExists(myArray, "X")


' =================================================
' isExists
' ^^^^^^^^
' �z���value�����݂��邩�`�F�b�N����
' �߂�l : [ TRUE:���݂��� | FALSE:���݂��Ȃ� ]
' =================================================
Function isExists(aryCheck, value)
	' �߂�l
	Dim returnValue
	returnValue = False
	
	' ���[�v�Ŏg�p����ϐ�
	Dim i, intMaxValue
	i = 0
	intMaxValue = Ubound(aryCheck)
	For i = 0 to intMaxValue
		if ( aryCheck(i) = value ) Then
			' ������������True������
			returnValue = True
		End If
	Next
	
	isExists = returnValue
	
End Function
