Option Explicit

Dim X
Set X = New TestClass

Dim modori
modori = X.DoTest("�H")
MsgBox modori

X.SubTest("subtest")

Set X = Nothing

Class TestClass
	Private Sub Class_Initialize
		MsgBox("TestClass �� ����܂����I")
	End Sub
	
	Public Function DoTest(pMsg)
		MsgBox pMsg
		DoTest = "["&pMsg&"]"
	End Function
	
	Public Sub SubTest(pMsg)
		MsgBox pMsg
	End Sub
	
	Private Sub Class_Terminate
		MsgBox("TestClass �� �j������܂����I")
	End Sub
End Class

