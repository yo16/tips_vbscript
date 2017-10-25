Option Explicit

Dim objWshShell
Set objWshShell = WScript.CreateObject("WScript.Shell")

Dim strStartupPath
strStartupPath = objWshShell.SpecialFolders("Startup")


MsgBox "スタートアップ:" & strStartupPath
