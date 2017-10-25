Option Explicit


Dim objFS
Set objFS = WScript.CreateObject("Scripting.FileSystemObject")

Dim objTS
Set objTS = objFS.OpenTextFile("sample.txt",1)

MsgBox objTS.ReadAll

objTS.Close
