Option Explicit
Dim objFS, objFolder, colFiles
Dim strFilesName, x
' FileSystemObject �I�u�W�F�N�g�𐶐�����
Set objFS = WScript.CreateObject("Scripting.FileSystemObject")
' �J�����g�t�H���_��Folder�I�u�W�F�N�g���擾����
Set objFolder = objFS.GetFolder(".")
' �J�����g�t�H���_�Ɋ܂܂�邷�ׂẴt�@�C�����擾����
Set colFiles = objFolder.Files
' ����\��
MsgBox colFiles.count
