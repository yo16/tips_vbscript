Option Explicit

' FunctionŒÄ‚Ño‚µ

Dim returnValue
returnValue = func_A(1, 2)
msgbox "FunctionŒÄ‚×‚½‚©‚È"&returnValue


' SubŒÄ‚Ño‚µ
Call sub_B("ŒÄ‚×‚é‚©‚È")




Function func_A(paramA, paramB)
	' –ß‚è’lİ’è
	func_A = paramA+paramB
End Function

Sub sub_B(param1)
	MsgBox(param1)
End Sub

