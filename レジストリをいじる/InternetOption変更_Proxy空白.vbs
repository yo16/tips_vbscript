Option Explicit


Dim objWshShell
Set objWshShell = WScript.CreateObject("WScript.Shell")

Dim regPath

' InternetOption�̐ݒ�
' �����\���X�N���v�g���g�p����
regPath = "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\AutoConfigURL"
objWshShell.RegDelete regPath


msgbox "ok"
