'�K�w�����ǂ��āA
'�t�@�C���E�t�H���_�ꗗ��
'�e�L�X�g�t�@�C���ɏo�͂���
Option Explicit

Dim fileName, overWrite
fileName = "fileTree.txt"
overWrite = True
' �t�@�C���o�̓t���O
Dim fileFlag
fileFlag = False



Dim objFS, objTS
Set objFS = CreateObject("Scripting.FileSystemObject")
Set objTS = objFS.CreateTextFile(fileName, overWrite)


' �h���b�v�����t�H���_�̖��O���o��
Dim objFolder, treeFolderName
Set objFolder = objFS.GetFolder(".")
treeFolderName = objFolder.Path


' �t�H���_�����t�@�C���ɏo��
printFolder treeFolderName, 0


Sub printFolder(folderName, floorNum)
	'�t���A��(�[��)�̕��A�^�u���o��
	Dim i
	For i = 0 to (floorNum-1)
		objTS.Write vbTab
	Next

	' �t�H���_�����o��
	objTS.WriteLine objFS.GetFileName(folderName)

	' �T�u�t�H���_���擾
	Dim objFolder
	Set objFolder = objFS.GetFolder(folderName)
	Dim objSubFolders, objSubFolder
	Set objSubFolders = objFolder.SubFolders

	' �T�u�t�H���_�̈ꗗ���o��
	For Each objSubFolder In objSubFolders
		printFolder folderName&"\"&objSubFolder.Name, floorNum+1
	Next

	If ( fileFlag ) Then
		' �t�@�C�����o��
		Dim folFiles
		Set folFiles = objFolder.Files
		Dim objFiles
		For Each objFiles in folFiles
			For i = 0 to floorNum
				objTS.Write vbTab
			Next
			objTS.WriteLine objFiles.Name
		Next
	End If
	
End Sub



