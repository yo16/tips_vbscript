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
	MsgBox filespec&"�́A����܂���ł�����H"
End If

If f.attributes and 16 Then
	MsgBox "�f�B���N�g���r�b�g�̓I���ł����B"
Else
	MsgBox "�f�B���N�g���r�b�g�̓I�t�ł����B"
End If

