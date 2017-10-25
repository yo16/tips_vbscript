Option Explicit

Dim objFS, objFolder
Set objFS = WScript.CreateObject("Scripting.FileSystemObject")
Set objFolder = objFS.GetFolder("a")

'Å´ìØÇ∂ÅB
'MsgBox objFolder.DateLastModified
MsgBox CStr(objFolder.DateLastModified)

