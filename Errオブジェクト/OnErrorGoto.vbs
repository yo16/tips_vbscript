Option Explicit

test1

msgbox "base"

Sub test1
'	On Error GoTo ErrorProcess1		' goto �ł��Ȃ��E�E�Ȃ�

	' �d���Ȃ��̂ŉ����
	On Error Resume Next
	Err.Raise(6)
	If ( Err.Number = 0 )Then
		MsgBox "no error"
	Else
		MsgBox "error!"
	End If
	On Error Goto 0
	exit sub

	ErrorProcess1:
		msgbox "error1"
End Sub