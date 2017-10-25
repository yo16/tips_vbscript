Option Explicit

Dim objFS
Set objFS = WScript.CreateObject("Scripting.FileSystemObject")
Dim objTS
Set objTS = objFS.OpenTextFile("sample2.txt", 1, false)


' Skipは、文字をスキップ
objTS.Skip(3)
Dim tmpLine
tmpLine = objTS.ReadLine
msgbox tmpLine

' SkipLineは、行をスキップ
objTS.SkipLine
tmpLine = objTS.ReadLine
msgbox tmpLine


