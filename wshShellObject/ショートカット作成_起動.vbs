' �V���[�g�J�b�g�쐬���Ǘ��҃��[�h�ŋN������
' 2016/3/15 yo16

Option Explicit

Dim objShell
Set objShell = CreateObject("Shell.Application")
objShell.ShellExecute "cscript.exe", "C:\zProgramming\VBScript\source\���K�\�[�X\wshShell�I�u�W�F�N�g\�V���[�g�J�b�g�쐬.vbs", "uac", "runas"
WScript.Quit
