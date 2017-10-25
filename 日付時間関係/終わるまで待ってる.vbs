Option Explicit

Dim WshShell
Set WshShell = WScript.CreateObject("WScript.Shell")

Dim fileName
fileName = "ç°èIÇÌÇ¡ÇΩÅI.txt"
Dim overWrite
overWrite = True


Dim modori
modori = WshShell.Run("C:\winnt\system32\cmd.exe",1,1)

Dim objFS,objTS
Set objFS = WScript.CreateObject("Scripting.FileSystemObject")
Set objTS = objFS.CreateTextFile(fileName,overWrite)

objTS.WriteLine time

objTS.Close


