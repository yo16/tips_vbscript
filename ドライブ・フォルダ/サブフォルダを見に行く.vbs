Option Explicit

Dim objFS
Set objFS = CreateObject("Scripting.FileSystemObject")

viewFolder("a")

Sub viewFolder(folderName)
	MsgBox folderName

	Dim objFolder
	Set objFolder = objFS.GetFolder(folderName)

	Dim objSubFolders,objSubFolder
	Set objSubFolders = objFolder.SubFolders

	For Each objSubFolder In objSubFolders
		viewFolder(folderName&"\"&objSubFolder.Name)
	Next

End Sub



