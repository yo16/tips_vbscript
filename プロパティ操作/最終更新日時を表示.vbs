Option Explicit

Dim fileName
fileName = "sample.txt"

Dim objFS,objFile
Set objFS = CreateObject("Scripting.FileSystemObject")
Set objFile = objFS.GetFile(fileName)

MsgBox objFile.DateLastModified,,fileName

