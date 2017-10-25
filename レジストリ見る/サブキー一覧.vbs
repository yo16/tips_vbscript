Option Explicit
' �C���|�[�g
Execute ReadFile("RegClass.vbs")
Execute ReadFile("RegClassCtl.vbs")


Dim WshShell
Set WshShell = WScript.CreateObject("WScript.Shell")

' ���W�X�g���G�N�X�|�[�g
Dim regKeyStr, regExpFile
regKeyStr = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
regExpFile = "export.txt"

'������:exe�t�@�C����
'������:�P���� �O���s��
'��O����:�P���I����҂� �O���҂����Ɏ������s
'�߂�l  :�O������I�� �P���ُ�I��
Dim rtn
rtn = WshShell.Run( "reg export " & regKeyStr & " " & regExpFile, 0, 1 )


' ���W�X�g����ǂ�ŁA���𐮗�
Dim regObj
Set regObj = GetRegClass(regExpFile, regKeyStr)

' ���W�X�g���I�u�W�F�N�g����A�����擾
Dim installPath_MemsONE, installPath_Mz
Dim subKeyCount
subKeyCount = regObj.GetSubKeysCount()
Dim valueStr
Dim i
For i = 0 to subKeyCount-1
	valueStr = regObj.GetSubKeyObjAt(i).GetValueByName( """DisplayName""" )
	If ( valueStr = """MemsONE""" ) Then
		installPath_MemsONE = regObj.GetSubKeyObjAt(i).GetValueByName( """InstallLocation""" )
		installPath_MemsONE = Replace(installPath_MemsONE, """", "")
		installPath_MemsONE = Replace(installPath_MemsONE, "\\", "\")
	ElseIf ( valueStr = """SMART PP 1.4""" ) Then
		installPath_Mz = regObj.GetSubKeyObjAt(i).GetValueByName( """InstallLocation""" )
		installPath_Mz = Replace(installPath_Mz, """", "")
		installPath_Mz = Replace(installPath_Mz, "\\", "\")
	End If
	
Next

msgbox installPath_MemsONE
msgbox installPath_Mz






' �O���t�@�C����ǂݍ��ށi�C���|�[�g�p�j
Function ReadFile(ByVal FileName)
	Const ForReading = 1
	
	Dim FileShell
	Set FileShell = WScript.CreateObject("Scripting.FileSystemObject")
	
	ReadFile = FileShell.OpenTextFile(FileName, ForReading, False).ReadAll()
End Function
