Option Explicit

Dim filespec
filespec = "a"

Dim fso, f
Set fso = CreateObject("Scripting.FileSystemObject")
If fso.FileExists(filespec) Then
	Set f = fso.GetFile(filespec)
ElseIf (fso.FolderExists(filespec)) Then
	Set f = fso.GetFolder(filespec)
Else
	MsgBox filespec&"は、ありませんでしたよ？"
End If

If f.attributes and 16 Then
	MsgBox "ディレクトリビットはオンでした。"
Else
	MsgBox "ディレクトリビットはオフでした。"
End If

