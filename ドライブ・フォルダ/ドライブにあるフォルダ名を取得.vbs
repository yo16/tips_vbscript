Option Explicit


Dim objFS, objFolder, colSubFolders
Dim strFoldersName, x

' FileSystemObject �I�u�W�F�N�g�𐶐�����
Set objFS = WScript.CreateObject("Scripting.FileSystemObject")

' Folder �I�u�W�F�N�g���擾����
Set objFolder = objFS.GetFolder(".")

' �T�u�t�H���_�� Folders �R���N�V�������擾����
Set colSubFolders = objFolder.SubFolders

' ���ׂẴT�u�t�H���_����strFoldersName�ɓ����
strFoldersName = ""
For Each x in colSubFolders
	strFoldersName = strFoldersName & x.Name & vbCRLF
Next

WScript.Echo strFoldersName



