Option Explicit

Dim myArray(2)
myArray(0) = "��"
myArray(1) = "��"
myArray(2) = "��"

'2�ڂ̈���""�́A��؂蕶���Ȃ����ĈӖ��B
'�����Ȃ��̏ꍇ�́A�X�y�[�X�ŋ�؂���B
MsgBox Join(myArray,"")
