Option Explicit
Dim objFS, objFolder, colSubFolders
Dim strFoldersName, strFoldersPath, x
' FileSystemObject �I�u�W�F�N�g�𐶐�����
Set objFS = WScript.CreateObject("Scripting.FileSystemObject")
' Folder �I�u�W�F�N�g���擾����
Set objFolder = objFS.GetFolder(".")
' �T�u�t�H���_�� Folders �R���N�V�������擾����
Set colSubFolders = objFolder.SubFolders
' ���ׂẴT�u�t�H���_����strFoldersName�ɓ����
strFoldersName = ""
strFoldersPath = ""
For Each x in colSubFolders
	' �t�@�C����
	strFoldersName = strFoldersName & x.Name & vbCRLF
	' �p�X
	strFoldersPath = strFoldersPath & x.Path & vbCrLf
Next
'WScript.Echo strFoldersName
WScript.Echo strFoldersPath
