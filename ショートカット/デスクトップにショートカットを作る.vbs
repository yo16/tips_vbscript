Option Explicit

Dim objWshShell, objShortcut
Dim strDesktopPath
' WshShell�I�u�W�F�N�g�𐶐�����
Set objWshShell = WScript.CreateObject("WScript.Shell")
' �f�X�N�g�b�v�̃t�H���_�����擾����
strDesktopPath = objWshShell.SpecialFolders("AllUsersDesktop")
msgbox strDesktopPath


Dim objFS
Set objFS = CreateObject("Scripting.FileSystemObject")
Dim objFile, objFileName, objFolder

Dim i, shortCutName
For i = 0 To WScript.Arguments.Count-1
	objFileName = WScript.Arguments(i)
	' ���̖��O�̂��̂��t�@�C�����t�H���_�����f
	If (objFS.FolderExists(objFileName) = -1) Then
		' �t�H���_
		Set objFolder = objFS.GetFolder(objFileName)
		shortCutName = objFolder.Name
	Else
		' �t�@�C��
		Set objFile = objFS.GetFile(objFileName)
		shortCutName = objFile.Name
	End If

	' WshShortcut�I�u�W�F�N�g�𐶐�����
	Set objShortcut = objWshShell.CreateShortcut(strDesktopPath & "\" & shortCutName & ".lnk")
	' �V���[�g�J�b�g�̃^�[�Q�b�g�t�@�C�����w�肷��
	objShortcut.TargetPath = objFileName
	' �V���[�g�J�b�g���쐬����
	objShortcut.Save
Next



