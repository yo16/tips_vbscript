Option Explicit

test1 1
test1 0
test1 2

msgbox "end"

Sub test1(num1)

	On Error Resume Next
	'Err.Raise(6)
	Dim num2
	num2 = 10/num1	' �������O�̂Ƃ��́A�[������
	If ( Err.Number = 0 )Then
		MsgBox "no error"
	Else
		MsgBox "error!("&Err.Number&")"
		Err.Clear		' �G���[�N���A
		On Error Goto 0		' Resume Next�̏�Ԃ����ɖ߂�
		Exit Sub
	End If
	
	On Error Goto 0		' Resume Next�̏�Ԃ����ɖ߂�
End Sub
