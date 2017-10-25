Option Explicit

Dim objFS

Set objFS = WScript.CreateObject("Scripting.FileSystemObject")


If (objFS.FolderExists("newFolder") = 0) Then
	objFS.CreateFolder "newFolder"
Else
	MsgBox "[newFolder]ÇÕä˘Ç…ë∂ç›ÇµÇ‹Ç∑ÅIÅI"
End If

