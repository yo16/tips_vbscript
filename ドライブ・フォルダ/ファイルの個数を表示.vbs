's099.vbs

Option Explicit
Dim objFS, objFolder, colFiles
' FileSystemObject �I�u�W�F�N�g�𐶐�����
Set objFS = WScript.CreateObject("Scripting.FileSystemObject")
' Folder �I�u�W�F�N�g���擾����
Set objFolder = objFS.GetFolder(".")
' Files �I�u�W�F�N�g���擾����
Set colFiles = objFolder.Files
' �t�@�C���̌���\������
WScript.Echo colFiles.Count
