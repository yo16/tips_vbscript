Option Explicit

Dim idx
idx = 1
Dim str
str = ""

Do While (idx < 10)
	str = str & idx
	idx = idx + 1
Loop

' ������ł��������ǁAExit Do�g���Ȃ�����Ȃ�ׂ��g���̂悻������
'While (idx < 10)
'	str = str & idx
'	idx = idx + 1
'Wend

msgbox str

