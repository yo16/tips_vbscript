' forでファイル一覧
' 2016/10/10 y.ikeda

Option Explicit

Dim objFS
Set objFS = CreateObject("Scripting.FileSystemObject")

Dim objFolder
Set objFolder = objFS.GetFolder(".")

' フォルダ一覧
Dim objSubFolder
For Each objSubFolder In objFolder.SubFolders
	MsgBox objSubFolder.Name
Next

' ファイル一覧
Dim objFile
For Each objFile In objFolder.Files
	MsgBox objFile.Name
Next
