Option Explicit


Dim objFS,objFolder
Set objFS = CreateObject("Scripting.FileSystemObject")
Set objFolder = objFS.GetFolder(".")
MsgBox objFolder.Path
Set objFolder = Nothing



' �ύX
Dim objWshShell
Set objWshShell = WScript.CreateObject("WScript.Shell")
objWshShell.CurrentDirectory = ".\a"



Set objFolder = objFS.GetFolder(".")
MsgBox objFolder.Path
