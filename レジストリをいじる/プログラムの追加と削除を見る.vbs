Option Explicit

' �v���O�����̒ǉ��ƍ폜�̖��O�ꗗ���o���Ă݂�


Dim objWshShell
Set objWshShell = WScript.CreateObject("WScript.Shell")

Dim regPath
regPath = "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"

Dim colEventSource
'colEventSource = objWshShell.RegRead(regPath)


'WScript.Echo colEventSource

Dim strSrc

For Each strSrc In objWshShell.RegRead(regPath)
    MsgBox "test"
    WScript.Echo strSrc
Next 



' ���߂��[�[�[
' �ꗗ���ق����̂ɁB�B2006/08/222006/08/222006/08/222006/08/222006/08/22
