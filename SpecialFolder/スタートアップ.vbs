Option Explicit

Dim objWshShell
Set objWshShell = WScript.CreateObject("WScript.Shell")

Dim strStartupPath
strStartupPath = objWshShell.SpecialFolders("Startup")


MsgBox "�X�^�[�g�A�b�v:" & strStartupPath
