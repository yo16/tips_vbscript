'�t�H���_�ɂ���t�@�C�����ꗗ��
'�e�L�X�g�t�@�C���ɏo�͂���

Option Explicit

Dim fileName
fileName = "fileNames.txt"
Dim overWrite
overWrite = True


Dim objFS, objFolder, colFiles, objTS
Dim strFilesName, x

Set objFS = WScript.CreateObject("Scripting.FileSystemObject")
Set objTS = objFS.CreateTextFile(fileName,overWrite)


Set objFolder = objFS.GetFolder(".")
Set colFiles = objFolder.Files

strFilesName = ""
For Each x in colFiles
	objTS.WriteLine x.Name
Next




