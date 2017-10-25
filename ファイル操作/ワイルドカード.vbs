Option Explicit

Dim objFS
Set objFS = Wscript.CreateObject("Scripting.FileSystemObject")

objFS.DeleteFile "aaaaa*.txt"
