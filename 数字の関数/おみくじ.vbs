Dim OmiValue, Response, MsgStr
Randomize   ' �����W�F�l���[�^���������B
OmiValue = Int((6 * Rnd) + 1)   ' 1 �` 6 �̃����_���Ȓl�𐶐��B


If ( OmiValue = 1 ) Then
	MsgStr = "����g��"
ElseIf ( OmiValue = 2 ) Then
	MsgStr = "�����g��"
ElseIf ( OmiValue = 3 ) Then
	MsgStr = "�����g��"
ElseIf ( OmiValue = 4 ) Then
	MsgStr = "�����g��"
ElseIf ( OmiValue = 5 ) Then
	MsgStr = "������"
Else
	MsgStr = "���勥��"
End If

MsgBox MsgStr, vbYes, "�����̉^��"



