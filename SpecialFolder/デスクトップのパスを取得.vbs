Option Explicit

Dim objWshShell
Set objWshShell = WScript.CreateObject("WScript.Shell")

Dim strDesktopPath
strDesktopPath = objWshShell.SpecialFolders("Desktop")


MsgBox "�f�X�N�g�b�v:"&strDesktopPath
