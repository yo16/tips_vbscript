Option Explicit

Dim objFS

' FileSystemObject �I�u�W�F�N�g�𐶐�����
Set objFS = WScript.CreateObject("Scripting.FileSystemObject")

' �t�@�C�������݂��邩�ǂ�����\������
' ���݂���ꍇ��True ���Ȃ��ꍇ��False
WScript.Echo objFS.FileExists("abc.txt")


