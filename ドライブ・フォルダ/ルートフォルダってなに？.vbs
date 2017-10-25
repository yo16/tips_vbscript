Option Explicit



Dim objFS,objDrive,objFolder

Set objFS = WScript.CreateObject("Scripting.FileSystemObject")

Set objDrive = objFS.GetDrive("D:")

Set objFolder = objDrive.RootFolder


Dim objSubFolders
Set objSubFolders = objFolder.SubFolders

Dim objSubFolder
For Each objSubFolder In objSubFolders
	MsgBox objSubFolder.Name
Next




