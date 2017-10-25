Option Explicit

Dim fileName
'********�o�̓t�@�C����********
fileName = "fileNames.txt"
Dim overWrite
overWrite = True


Dim objFS, objTS
Set objFS = WScript.CreateObject("Scripting.FileSystemObject")
Set objTS = objFS.CreateTextFile(fileName,overWrite)

Dim objFolder
Set objFolder = objFS.GetFolder(".")

'�t�@�C���ꗗ���쐬
objTS.WriteLine "** �t�@�C���ꗗ **"
Dim colFiles
Set colFiles = objFolder.Files
Dim x
For Each x in colFiles
	If ((x.Name <> fileName) and (x.Name <> WScript.ScriptName)) Then
		objTS.WriteLine x.Name
	End If
Next

'�t�H���_�ꗗ���쐬
objTS.WriteLine "** �t�H���_�ꗗ **"
Dim colSubFolders
Set colSubFolders = objFolder.SubFolders
For Each x in colSubFolders
	objTS.WriteLine x.Name
Next


objTS.Close

