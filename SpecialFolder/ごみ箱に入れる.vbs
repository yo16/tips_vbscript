Option Explicit

Dim objFS
Set objFS = WScript.CreateObject("Scripting.FileSystemObject")

'Dim objFile
'Set objFile = objFS.GetFile("sample.txt")

Dim objWshShell
Set objWshShell = WScript.CreateObject("WScript.Shell")

msgbox objWshShell.SpecialFolders("Desktop")&"\���ݔ�\sample.txt"
'objFS.MoveFile "sample.txt",objWshShell.SpecialFolders("Desktop")&"\���ݔ�\sample.txt"
objFS.MoveFile "sample.txt",objWshShell.SpecialFolders("Desktop")&"\���ݔ�\"
