Option Explicit
' �t�@�C������like��������
' 2006/12/22 ikeda

Dim path1

path1 = "C:\900_Programming\VBScript\source\���K�\�[�X\�����񑀍�\������*"


Dim objFS, objFolder, colFiles
Dim strFilesName, x
' FileSystemObject �I�u�W�F�N�g�𐶐�����
Set objFS = WScript.CreateObject("Scripting.FileSystemObject")
' �J�����g�t�H���_��Folder�I�u�W�F�N�g���擾����
Set objFolder = objFS.GetFolder(".")
' �J�����g�t�H���_�Ɋ܂܂�邷�ׂẴt�@�C�����擾����
Set colFiles = objFolder.Files
' �X�̃t�@�C�����𕶎���ɒǉ�����
strFilesName = ""
For Each x in colFiles
	strFilesName = strFilesName & x.Name & vbCRLF
Next
' ���ʂ�\������
WScript.Echo strFilesName
