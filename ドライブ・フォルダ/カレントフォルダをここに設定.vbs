' �J�����g�t�H���_���A���̃X�N���v�g���u���Ă���t�H���_�֕ύX
' 2017/03/03 (c) yo16

Dim objFS,objFolder
Set objFS = CreateObject("Scripting.FileSystemObject")
Set objFolder = objFS.GetFolder(".")
msgbox objFolder.Path	' �m�F�p



' �X�N���v�g�̃p�X
msgbox WScript.ScriptFullName	' �m�F�p

' �t�H���_�p�X�𒊏o
Dim scriptDir
scriptDir = objFS.getParentFolderName( WScript.ScriptFullName )
msgbox scriptDir	' �m�F�p



' �J�����g�t�H���_��ύX
Dim objWshShell
Set objWshShell = WScript.CreateObject("WScript.Shell")
objWshShell.CurrentDirectory = scriptDir




Set objFolder = objFS.GetFolder(".")
msgbox objFolder.Path	' �m�F�p

