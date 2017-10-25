Option Explicit

Dim testStr
testStr = "x:\aaa\abc.txt"

Dim objFS
Set objFS = CreateObject("Scripting.FileSystemObject")

MsgBox testStr,,"‚à‚Æ‚à‚Æ"
' x:\aaa\abc.txt
MsgBox objFS.GetBaseName(testStr),,"BaseName"
' abc
MsgBox objFS.GetExtensionName(testStr),,"ExtensionName"
' txt
MsgBox objFS.GetFileName(testStr),,"FileName"
' abc.txt
MsgBox objFS.GetAbsolutePathName(testStr),,"GetAbsolutePathName"
' X:\aaa\abc.txt
' ‚È‚º‚©‘å•¶Žš‚É‚È‚é
