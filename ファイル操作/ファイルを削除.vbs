Option Explicit

' 2008/01/17 ikeda
' �t�@�C�����폜
Dim objFS
Set objFS = CreateObject("Scripting.FileSystemObject")

If ( Not objFS.FileExists("�폜") )Then
	msgBox "�e�X�g�p�t�@�C�� �폜.txt ������܂���"
	WScript.Quit
End If

'objFS.deleteFile "�폜.txt"



' �ǂݎ���p���폜���邩�m�F
' �� �ǂݎ���p��NG�I

' �ǂݎ���p�΍�
Dim objFile
Set objFile = objFS.GetFile( "�폜.txt" )

' File�I�u�W�F�N�g��Attributes�v���p�e�B��ύX����
' �ǂݎ���p�́A2�r�b�g��
If ( objFile.Attributes and 1 ) Then
	' �ǂݎ���p�t���O�������Ă�����|��
	objFile.Attributes = objFile.Attributes - 1
End If

objFS.deleteFile "�폜.txt"
