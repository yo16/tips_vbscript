Option Explicit


Dim fileName
fileName = "empty.txt"
Dim overWrite
overWrite = True

Dim objFS,objTS
Set objFS = WScript.CreateObject("Scripting.FileSystemObject")
Set objTS = objFS.CreateTextFile(fileName,overWrite)

objTS.Close


