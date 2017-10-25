
Class SotoClass
	Private Sub Class_Initialize
		MsgBox("SotoClass ‚ª ì‚ç‚ê‚Ü‚µ‚½I")
	End Sub
	Public Function DoTest(pMsg)
		MsgBox pMsg
		DoTest = "["&pMsg&"]"
	End Function
	Private Sub Class_Terminate
		MsgBox("SotoClass ‚ª ”jŠü‚³‚ê‚Ü‚µ‚½I")
	End Sub
End Class

