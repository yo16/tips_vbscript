' �����𓾂�
' 2005/11/10
' 2006/04/27 memo Drag&Drop����ƁA�t���p�X����P�����ɂȂ��ċN�������B
'                 (0)���A�������Bc�̂悤�Ɏ������g�̃t�@�C�����ł͂Ȃ��B

Option Explicit

Dim objArgs, I
Set objArgs = WScript.Arguments

WScript.Echo "�����̐�:" & objArgs.Count

For I = 0 to objArgs.Count - 1
	WScript.Echo "���������I:" & objArgs(I)
Next


