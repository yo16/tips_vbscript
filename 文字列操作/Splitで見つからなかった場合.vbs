Option Explicit
' Split�Ō�����Ȃ������ꍇ�A(0)�ɓ���̂��H

Dim aryFound
aryFound = Split("abcde", "/")

MsgBox aryFound(0)		' abcde
MsgBox UBound(aryFound)	' 0
' (0)�ɂ͂���AUBound��0�ɂȂ�
