Option Explicit


Dim objWshShell
Set objWshShell = WScript.CreateObject("WScript.Shell")

Dim regPath

' InternetOption�̐ݒ�
' �����\���X�N���v�g���g�p����
regPath = "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\AutoConfigURL"
objWshShell.RegWrite regPath,"http://net.XXXX.co.jp/net/proxy.pac","REG_SZ"


msgbox "ok"
