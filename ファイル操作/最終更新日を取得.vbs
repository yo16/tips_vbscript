Option Explicit

Dim objFS, objFolder
Set objFS = WScript.CreateObject("Scripting.FileSystemObject")
Set objFolder = objFS.GetFolder("a")

'�������B
'MsgBox objFolder.DateLastModified
MsgBox CStr(objFolder.DateLastModified)

