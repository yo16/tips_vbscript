Option Explicit

Dim objFS
Set objFS = CreateObject("Scripting.FileSystemObject")

Dim newFolderName
newFolderName = "newFolder"




viewFolder("a")



Dim objWshShell
Set objWshShell = WScript.CreateObject("WScript.Shell")
Dim modori
modori = objWshShell.PopUp(newFolderName,1,"title",0)


Sub viewFolder(folderName)
'	MsgBox folderName

	If (objFS.FolderExists(folderName&"\"&newFolderName) = 0) Then
		objFS.CreateFolder folderName&"\"&newFolderName
	End If

	Dim objFolder
	Set objFolder = objFS.GetFolder(folderName)

	Dim objSubFolders,objSubFolder
	Set objSubFolders = objFolder.SubFolders

	For Each objSubFolder In objSubFolders
		If (objSubFolder.Name <> newFolderName) Then
			viewFolder(folderName&"\"&objSubFolder.Name)
		End If
	Next

End Sub



