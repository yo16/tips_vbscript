Dim WSHShell,intErrCode
Set WSHShell = WScript.CreateObject("WScript.Shell")
intErrCode=WSHShell.Run("D:\wsh\main\main.wsf",1,True)

Select Case intErrCode
	Case -1 MsgBox "�_�C�A���O�͎����I�ɕ����܂����B"
	Case -2 MsgBox	"�G���[�Ȃ񂾂ȁI"
	Case vbYes MsgBox "�u�͂��v�������܂����B:  " & vbYes
	Case vbNo MsgBox "�u�������v�������܂����B:  " & vbNo
	Case vbCancel MsgBox "�u�L�����Z���v�������܂����B:  " & vbCancel
End Select

