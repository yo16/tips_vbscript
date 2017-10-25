Option Explicit

Dim objFS

Set objFS = WScript.CreateObject("Scripting.FileSystemObject")


' ２個直列に、一気に作ろうとするとエラーになる
If (objFS.FolderExists("test\newFolder") = 0) Then
	objFS.CreateFolder "test\newFolder"
Else
	MsgBox "[newFolder]は既に存在します！！"
End If

