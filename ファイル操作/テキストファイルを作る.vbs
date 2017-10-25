Option Explicit
' ÉtÉ@ÉCÉãçÏê¨

Dim fileName
fileName = "abc.txt"
Dim overWrite
overWrite = True

Dim objFS,objTS
Set objFS = WScript.CreateObject("Scripting.FileSystemObject")
Set objTS = objFS.CreateTextFile(fileName,overWrite)

objTS.WriteLine "abccba"
objTS.Write "aaa"
objTS.Write "bbb"
objTS.Write vbCrLf

objTS.Close

Set objTS = Nothing
Set objFS = Nothing
