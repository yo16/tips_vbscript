' �t�H���_�폜
' 2015/4/24

Option Explicit

Dim objFs
Set objFs = CreateObject("Scripting.FileSystemObject")

Dim targetFolderName
targetFolderName = "delTest"

' �e�X�g�p�t�H���_���쐬
If( Not objFs.FolderExists(targetFolderName) )Then
	objFs.CreateFolder targetFolderName
End If
If( Not objFs.FolderExists(targetFolderName & "\subFolder1") )Then
	objFs.CreateFolder targetFolderName & "\subFolder1"
End If

MsgBox targetFolderName & "���폜���܂��I"


objFs.DeleteFolder targetFolderName
' �� ���Ƀt�@�C����t�H���_�������Ă��A�������ɍ폜�����

