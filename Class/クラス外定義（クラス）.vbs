
Class SotoClass
	Private Sub Class_Initialize
		MsgBox("SotoClass �� ����܂����I")
	End Sub
	Public Function DoTest(pMsg)
		MsgBox pMsg
		DoTest = "["&pMsg&"]"
	End Function
	Private Sub Class_Terminate
		MsgBox("SotoClass �� �j������܂����I")
	End Sub
End Class

