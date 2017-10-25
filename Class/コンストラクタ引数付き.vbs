' できない！2006/10/11

Option Explicit

Dim X
Set X = New TestClass("test")
' ！！！渡せません！

X.SubTest()

Set X = Nothing





Class TestClass
	Dim m_param
	
	Private Sub Class_Initialize(param)
		MsgBox("TestClass が 作られました！")
		m_param = param
	End Sub
	
	Public Sub SubTest()
		MsgBox m_param
	End Sub
	
	Private Sub Class_Terminate
		MsgBox("TestClass が 破棄されました！")
	End Sub
End Class

