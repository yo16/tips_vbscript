'�t�H���_�ɂ���t�H���_���ꗗ��
'�e�L�X�g�t�@�C���ɏo�͂���

Option Explicit

Dim fileName
fileName = "fileNames.txt"
Dim overWrite
overWrite = True


Dim objFS, objFolder, colSubFolders, objTS
Dim strFilesName, x

Set objFS = WScript.CreateObject("Scripting.FileSystemObject")
Set objTS = objFS.CreateTextFile(fileName,overWrite)


Set objFolder = objFS.GetFolder(".")
Set colSubFolders = objFolder.SubFolders


For Each x in colFiles
	objTS.WriteLine x.Name
Next


