Option Explicit

Dim objFS

Set objFS = WScript.CreateObject("Scripting.FileSystemObject")


' �Q����ɁA��C�ɍ�낤�Ƃ���ƃG���[�ɂȂ�
If (objFS.FolderExists("test\newFolder") = 0) Then
	objFS.CreateFolder "test\newFolder"
Else
	MsgBox "[newFolder]�͊��ɑ��݂��܂��I�I"
End If

