'-----------------------------------
' �ǂݎ���p�v���p�e�B��ύX����
'-----------------------------------
Option Explicit

Dim FileName
FileName = "readonly.txt"

Dim objFS
Set objFS = CreateObject("Scripting.FileSystemObject")
Dim objFile
Set objFile = objFS.GetFile( FileName )

' File�I�u�W�F�N�g��Attributes�v���p�e�B��ύX����
' �ǂݎ���p�́A2�r�b�g��
If ( objFile.Attributes and 1 ) Then
	' �ǂݎ���p�t���O�������Ă�����|��
	objFile.Attributes = objFile.Attributes - 1
Else
	' �ǂݎ���p�t���O���|��Ă����痧�Ă�
	objFile.Attributes = objFile.Attributes + 1
End If

