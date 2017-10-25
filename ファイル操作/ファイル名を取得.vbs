Option Explicit

Dim testStr
testStr = "x:\aaa\abc.txt"

Dim objFS
Set objFS = CreateObject("Scripting.FileSystemObject")

MsgBox testStr,,"もともと"
' x:\aaa\abc.txt
MsgBox objFS.GetBaseName(testStr),,"BaseName"
' abc
MsgBox objFS.GetExtensionName(testStr),,"ExtensionName"
' txt
MsgBox objFS.GetFileName(testStr),,"FileName"
' abc.txt
MsgBox objFS.GetAbsolutePathName(testStr),,"GetAbsolutePathName"
' X:\aaa\abc.txt
' なぜか大文字になる
