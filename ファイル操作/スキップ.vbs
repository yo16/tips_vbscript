Option Explicit

Dim objFS
Set objFS = WScript.CreateObject("Scripting.FileSystemObject")
Dim objTS
Set objTS = objFS.OpenTextFile("sample2.txt", 1, false)


' Skip�́A�������X�L�b�v
objTS.Skip(3)
Dim tmpLine
tmpLine = objTS.ReadLine
msgbox tmpLine

' SkipLine�́A�s���X�L�b�v
objTS.SkipLine
tmpLine = objTS.ReadLine
msgbox tmpLine


