' �ł��Ȃ��I2006/10/11

Option Explicit

Dim X
Set X = New TestClass("test")
' �I�I�I�n���܂���I

X.SubTest()

Set X = Nothing





Class TestClass
	Dim m_param
	
	Private Sub Class_Initialize(param)
		MsgBox("TestClass �� ����܂����I")
		m_param = param
	End Sub
	
	Public Sub SubTest()
		MsgBox m_param
	End Sub
	
	Private Sub Class_Terminate
		MsgBox("TestClass �� �j������܂����I")
	End Sub
End Class

