' �����̃t�@�C�������擾

' �t�@�C��������
MsgBox WScript.ScriptName

' �t�@�C�������܂ރt���p�X
MsgBox WScript.ScriptFullName

' �t�H���_�������i\���܂߂Ȃ����}�C�i�X1�̕��j
MsgBox Left( _
	WScript.ScriptFullName, _
	Len(WScript.ScriptFullName) - Len(WScript.ScriptName) - 1 _
)
