Option Explicit


Dim objWshShell
Set objWshShell = WScript.CreateObject("WScript.Shell")

Dim regPath

' ログファイル
regPath = "HKCU\Software\HSTools\IPMsg17777\LogCheck"
objWshShell.RegWrite regPath,1,"REG_DWORD"

' ログを書く
regPath = "HKCU\Software\HSTools\IPMsg17777\LogFile"
objWshShell.RegWrite regPath,"c:\Program Files\ipm\ipmsg.log"


'MsgBox "終了"


