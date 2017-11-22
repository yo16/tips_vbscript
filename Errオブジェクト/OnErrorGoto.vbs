Option Explicit

test1 1
test1 0
test1 2

msgbox "end"

Sub test1(num1)

	On Error Resume Next
	'Err.Raise(6)
	Dim num2
	num2 = 10/num1	' 引数が０のときは、ゼロ割り
	If ( Err.Number = 0 )Then
		MsgBox "no error"
	Else
		MsgBox "error!("&Err.Number&")"
		Err.Clear		' エラークリア
		On Error Goto 0		' Resume Nextの状態を元に戻す
		Exit Sub
	End If
	
	On Error Goto 0		' Resume Nextの状態を元に戻す
End Sub
