Option Explicit

Dim objFS

' FileSystemObject �I�u�W�F�N�g�𐶐�����
Set objFS = WScript.CreateObject("Scripting.FileSystemObject")

' �t�H���_�����݂��邩�ǂ�����\������
'���݂���ꍇ��		-1
'���݂��Ȃ��ꍇ��		0
'	��Ԃ�
WScript.Echo objFS.FolderExists("c:\WINDOWS")


