Option Explicit

Dim X
Set X = New TestClass

Dim modori
modori = X.DoTest("ÅH")
MsgBox modori

Set X = Nothing

Class List
	Private nextPointer
	Private prevPointer
	private listValue

'* Initialize,Terminate *
	Private Sub Class_Initialize
'		MsgBox("TestClass Ç™ çÏÇÁÇÍÇ‹ÇµÇΩÅI")
	End Sub
	Private Sub Class_Terminate
'		MsgBox("TestClass Ç™ îjä¸Ç≥ÇÍÇ‹ÇµÇΩÅI")
	End Sub

'* Property [next] *
	Private Property Get next()
		next = nextPointer
	End Property
	Private Property Let next(p_next)
		nextPointer = p_next
	End Property

'* Property [prev] *
	Private Property Get prev()
		prev = prevPointer
	End Property
	Private Property Let prev(p_prev)
		prevPointer = p_prev
	End Property


End Class

