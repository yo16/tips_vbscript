Option Explicit

' �G�ۂ̃C���X�g�[���p�X


Dim objWshShell
Set objWshShell = WScript.CreateObject("WScript.Shell")

Dim regPath
'regPath = "HKLM\SOFTWARE\Classes\Applications\Hidemaru.exe\shell\edit\command\"
regPath = "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Hidemaru\UninstallString"

Dim hidePath
hidePath = objWshShell.RegRead(regPath)

' �A���C���X�g�[���p������ " /R" �������ɓ����Ă�̂ŏ���
Dim foundPos
foundPos = InStrRev(hidePath, " /R")
Dim hidePath2
hidePath2 = Left(hidePath, foundPos-1)


msgbox "[" & hidePath & "]"
msgbox "[" & hidePath2 & "]"

