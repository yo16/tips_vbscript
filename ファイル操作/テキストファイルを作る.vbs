Option Explicit
' ファイル作成

Dim fileName
fileName = "abc.txt"
Dim overWrite
overWrite = True

Dim objFS,objTS
Set objFS = WScript.CreateObject("Scripting.FileSystemObject")
' 作る場合は下記のいずれでも同じ
'Set objTS = objFS.CreateTextFile(fileName,overWrite)
'Set objTS = objFS.OpenTextFile(fileName, 2, true)	' 1:ForReading | 2:ForWriting | 8:ForAppending
Set objTS = objFS.OpenTextFile(fileName, 2, 1)

objTS.WriteLine "abccba"
objTS.Write "aaa"
objTS.Write "bbb"
objTS.Write vbCrLf

objTS.Close

Set objTS = Nothing
Set objFS = Nothing
