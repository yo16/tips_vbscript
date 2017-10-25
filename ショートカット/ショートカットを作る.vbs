Option Explicit
Dim objWshShell, objShortcut
Dim strDesktopPath
Set objWshShell = WScript.CreateObject("WScript.Shell")

Dim oracleHomePath
oracleHomePath = GetOracleHome()

strDesktopPath = objWshShell.SpecialFolders("Desktop")

Set objShortcut = objWshShell.CreateShortcut(strDesktopPath & "\�\�[�X�Ǘ�.lnk")

' �V���[�g�J�b�g�̃^�[�Q�b�g�t�@�C�����w�肷��
objShortcut.TargetPath = oracleHomePath & "\Bin\ifrun60.EXE"
' �V���[�g�J�b�g�ɓn���������w�肷��
objShortcut.Arguments = "c:\prism\form\�t�@�C���Ǘ�.fmx nova01/nova01@smtap.world"
' �V���[�g�J�b�g���쐬����
objShortcut.Save





Function GetOracleHome()
'---------------------------------------------------
' ���W�X�g�����Q�Ƃ��āAORACLE_HOME�̃p�X��Ԃ��֐�
'---------------------------------------------------
	Dim objWshShell
	Dim RegData

	Set objWshShell = WScript.CreateObject ("WScript.Shell")
	RegData = "HKLM\Software\ORACLE\ORACLE_HOME"
	GetOracleHome = objWshShell.RegRead(RegData)

End Function
