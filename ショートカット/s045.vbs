Option Explicit
Dim objWshShell, objShortcut
Dim strDesktopPath
' WshShell�I�u�W�F�N�g�𐶐�����
Set objWshShell = WScript.CreateObject("WScript.Shell")
' �f�X�N�g�b�v�̃t�H���_�����擾����
strDesktopPath = objWshShell.SpecialFolders("Desktop")
' WshShortcut�I�u�W�F�N�g�𐶐�����
Set objShortcut = objWshShell.CreateShortcut(strDesktopPath & "\�����̕\.lnk")
' �V���[�g�J�b�g�̃^�[�Q�b�g�t�@�C�����w�肷��
objShortcut.TargetPath = "c:\program files\microsoft office\office\excel.exe"
' �V���[�g�J�b�g�ɓn���������w�肷��
objShortcut.Arguments = "c:\home\wsh\ch02\s1.xls c:\home\wsh\ch02\s2.xls"
' �V���[�g�J�b�g���쐬����
objShortcut.Save
