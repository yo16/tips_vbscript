Option Explicit
Dim objWshShell, objShortcut
Dim strDesktopPath
' WshShell�I�u�W�F�N�g�𐶐�����
Set objWshShell = WScript.CreateObject("WScript.Shell")
' WshShortcut�I�u�W�F�N�g�𐶐�����
Set objShortcut = objWshShell.CreateShortcut("test.lnk")
' �V���[�g�J�b�g�̃^�[�Q�b�g�t�@�C�����w�肷��
objShortcut.TargetPath = "c:\test"
' �V���[�g�J�b�g���쐬����
objShortcut.Save
