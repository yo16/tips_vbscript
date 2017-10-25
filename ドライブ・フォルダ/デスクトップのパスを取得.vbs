'2001/12/17

Option Explicit

Dim objWshShell, objShortcut
Dim strDesktopPath
' WshShellオブジェクトを生成する
Set objWshShell = WScript.CreateObject("WScript.Shell")
' デスクトップのフォルダ名を取得する
strDesktopPath = objWshShell.SpecialFolders("Desktop")

MsgBox strDesktopPath
