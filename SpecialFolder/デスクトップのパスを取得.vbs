Option Explicit

Dim objWshShell
Set objWshShell = WScript.CreateObject("WScript.Shell")

Dim strDesktopPath
strDesktopPath = objWshShell.SpecialFolders("Desktop")


MsgBox "デスクトップ:"&strDesktopPath
