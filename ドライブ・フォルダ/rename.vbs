'************************
' フォルダをリネームする
' rename_A <-> rename_B
'************************
Option Explicit

Dim objFS
Set objFS = WScript.CreateObject("Scripting.FileSystemObject")

Dim objFolder

If ( objFS.FolderExists(".\rename_A") ) Then
	Set objFolder = objFS.GetFolder(".\rename_A")
	objFolder.Name = "rename_B"
Else
	Set objFolder = objFS.GetFolder(".\rename_B")
	objFolder.Name = "rename_A"
End If

