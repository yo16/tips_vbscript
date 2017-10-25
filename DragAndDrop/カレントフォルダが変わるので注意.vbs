Option Explicit

Dim objFS,objFolder
Set objFS = CreateObject("Scripting.FileSystemObject")
Set objFolder = objFS.GetFolder(".")
Msgbox "D&Dのときは、カレントフォルダが変わります" & vbCrLf & objFolder.Path



' そしてその対策
dim objShell
set objShell = CreateObject("WScript.Shell")
objShell.CurrentDirectory = objFS.GetParentFolderName(WScript.ScriptFullName)


Set objFolder = objFS.GetFolder(".")
Msgbox "２度目は？" & vbCrLf & objFolder.Path
