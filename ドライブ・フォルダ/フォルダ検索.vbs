'--------------------------------------------------------
'�t�H���_���������A��������t�@�C���֏o�͂���
'--------------------------------------------------------
Option Explicit

'�T��������
Dim findStr
findStr = "a"

Dim fileName, overWrite
fileName = "findFile.txt"
overWrite = True



Dim objFS, objTS
Set objFS = CreateObject("Scripting.FileSystemObject")
Set objTS = objFS.CreateTextFile(fileName, overWrite)


' �h���b�v�����t�H���_�̖��O���o��
Dim objFolder, treeFolderName
Set objFolder = objFS.GetFolder(".")
treeFolderName = objFolder.Path


' �t�H���_�����t�@�C���ɏo��
printFolder treeFolderName


Sub printFolder(folderName)
	' �������e�Ɉ�v���邩�m�F
	Dim pos
	pos = Instr( objFS.GetFileName(folderName), findStr )
	If ( pos <> 0 ) Then
		' �t�H���_�����o��
		objTS.WriteLine objFS.GetAbsolutePathName(folderName)
	End If
	
	' �t�H���_�I�u�W�F�N�g���擾
	Dim objFolder
	Set objFolder = objFS.GetFolder(folderName)
	
	' �T�u�t�H���_���擾
	Dim objSubFolders, objSubFolder
	Set objSubFolders = objFolder.SubFolders
	
	' �T�u�t�H���_�̈ꗗ���o��
	For Each objSubFolder In objSubFolders
		printFolder folderName&"\"&objSubFolder.Name
	Next
	
End Sub



