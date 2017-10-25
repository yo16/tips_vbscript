Option Explicit


Dim objWshShell
Set objWshShell = WScript.CreateObject("WScript.Shell")

Dim regPath

' InternetOptionの設定
' 自動構成スクリプトを使用する
regPath = "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\AutoConfigURL"
objWshShell.RegDelete regPath


msgbox "ok"
