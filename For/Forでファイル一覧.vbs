' for�Ńt�@�C���ꗗ
' 2016/10/10 y.ikeda

Option Explicit

Dim objFS
Set objFS = CreateObject("Scripting.FileSystemObject")

Dim objFolder
Set objFolder = objFS.GetFolder(".")

' �t�H���_�ꗗ
Dim objSubFolder
For Each objSubFolder In objFolder.SubFolders
	MsgBox objSubFolder.Name
Next

' �t�@�C���ꗗ
Dim objFile
For Each objFile In objFolder.Files
	MsgBox objFile.Name
Next
