Option Explicit


Dim a
Dim b

a = 123
b = 456

MsgBox "a = " & a & vbCrLf & "b = " & b,,"初期値"


'参照呼び
call byRefSub(a,b)

MsgBox "a*10 = " & a & vbCrLf & "b*10 = " & b,,"参照呼び"


a = 123
b = 456

'値呼び
call byValSub((a),(b))

MsgBox "a*10 = " & a & vbCrLf & "b*10 = " & b,,"値呼び"





Sub byRefSub(byRef param1,param2)		'param2も参照渡しになる
	param1 = param1 * 10
	param2 = param2 * 10
End Sub

Sub byValSub(byVal param1,param2)
	param1 = param1 * 10
	param2 = param2 * 10
End Sub

