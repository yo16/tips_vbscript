Option Explicit

Dim fileName		: fileName = "a.txt"

Dim objShell		: Set objShell = WScript.CreateObject("WScript.Shell")
Dim objShellAp		: Set objShellAp = WScript.CreateObject("Shell.Application")

' カレントフォルダにある、
Dim objFolder		: Set objFolder = objShellAp.Namespace(objShell.CurrentDirectory)
' a.txtの、
Dim objFolderItem	: Set objFolderItem = objFolder.ParseName(fileName)
' 更新日時を変更
objFolderItem.ModifyDate = CDate("2011/11/11 11:11:11")

