'2001/12/17

Option Explicit

Dim objWshShell, objShortcut
Dim strDesktopPath
' WshShell�I�u�W�F�N�g�𐶐�����
Set objWshShell = WScript.CreateObject("WScript.Shell")
' �f�X�N�g�b�v�̃t�H���_�����擾����
strDesktopPath = objWshShell.SpecialFolders("Desktop")

MsgBox strDesktopPath
