Option Explicit


Dim a
Dim b

a = 123
b = 456

MsgBox "a = " & a & vbCrLf & "b = " & b,,"�����l"


'�Q�ƌĂ�
call byRefSub(a,b)

MsgBox "a*10 = " & a & vbCrLf & "b*10 = " & b,,"�Q�ƌĂ�"


a = 123
b = 456

'�l�Ă�
call byValSub((a),(b))

MsgBox "a*10 = " & a & vbCrLf & "b*10 = " & b,,"�l�Ă�"





Sub byRefSub(byRef param1,param2)		'param2���Q�Ɠn���ɂȂ�
	param1 = param1 * 10
	param2 = param2 * 10
End Sub

Sub byValSub(byVal param1,param2)
	param1 = param1 * 10
	param2 = param2 * 10
End Sub

