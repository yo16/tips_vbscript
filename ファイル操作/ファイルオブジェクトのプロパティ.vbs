Option Explicit
' �t�@�C���I�u�W�F�N�g�̃v���p�e�B

Dim objFs
Set objFs = CreateObject("Scripting.FileSystemObject")

Dim objFile
Set objFile = objFs.GetFile("a.txt")

MsgBox objFile.Name ,vbOkOnly,"�t�@�C����"
MsgBox objFile.Path ,vbOkOnly,"�t�@�C���p�X"
MsgBox objFile.ParentFolder ,vbOkOnly,"ParentFolder"
MsgBox objFile.DateCreated ,vbOkOnly,"�쐬����"
MsgBox objFile.DateLastModified ,vbOkOnly,"�X�V����"

Set objFile = Nothing
Set objFs = Nothing
