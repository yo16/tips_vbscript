Option Explicit


Dim objFS
Set objFS = WScript.CreateObject("Scripting.FileSystemObject")

objFS.CopyFile "abc.txt",".\フォルダ指定\"
WScript.Quit

