Option Explicit
On Error Resume Next

'�͂���
If Err Then
	MsgBox "�G���[�I",,"�͂���"
Else
	MsgBox "����B",,"�͂���"
End If

'�G���[���N�����Ă݂�
Err.Raise(10)
If Err Then
	MsgBox "�G���[�I",,"�G���[���N�����Ă݂�"
Else
	MsgBox "����B",,"�G���[���N�����Ă݂�"
End If

'�ق��Ƃ��Ă݂�
If Err Then
	MsgBox "�G���[�I",,"�ق��Ƃ��Ă݂�"
Else
	MsgBox "����B",,"�ق��Ƃ��Ă݂�"
End If

'�N���A���Ă݂�
Err.Clear
If Err Then
	MsgBox "�G���[�I",,"�N���A���Ă݂�"
Else
	MsgBox "����B",,"�N���A���Ă݂�"
End If
