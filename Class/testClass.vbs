Option Explicit

Dim X
Set X = New TestClass

Dim modori
modori = X.DoTest("ÅH")
MsgBox modori

X.SubTest("subtest")

Set X = Nothing

Class TestClass
	Private Sub Class_Initialize
		MsgBox("TestClass Ç™ çÏÇÁÇÍÇ‹ÇµÇΩÅI")
	End Sub
	
	Public Function DoTest(pMsg)
		MsgBox pMsg
		DoTest = "["&pMsg&"]"
	End Function
	
	Public Sub SubTest(pMsg)
		MsgBox pMsg
	End Sub
	
	Private Sub Class_Terminate
		MsgBox("TestClass Ç™ îjä¸Ç≥ÇÍÇ‹ÇµÇΩÅI")
	End Sub
End Class

