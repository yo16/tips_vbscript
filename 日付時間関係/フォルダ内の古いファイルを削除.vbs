Option Explicit

Dim objFS
Set objFS = WScript.CreateObject("Scripting.FileSystemObject")

Dim objFolder
Set objFolder = objFS.GetFolder("�폜�e�X�g�p�t�H���_")

Dim colFiles
Set colFiles = objFolder.Files

Dim x,deleteCount
deleteCount = 0
For Each x in colFiles
	If deleteOldFile(x,5) Then deleteCount = deleteCount + 1
Next

MsgBox deleteCount & "�̃t�@�C�����폜���܂����I",,"�� ���ʕ� ��"

'**************************************************
'�֐�[deleteOldFile]
'objFile	:�폜����Ώۂ̃t�@�C���I�u�W�F�N�g
'stockDay	:�ۑ��������(���ɂ��P��)
'�߂�l		:�폜�����ꍇ��True
'			 �폜���Ȃ������ꍇ��False
'
'[fileName]�̍ŏI�X�V����[stockDay]���ȏ�O�̏ꍇ��
'�폜����֐��B
'**************************************************
Function deleteOldFile(objFile,stockDay)
	Dim dateDifference
	dateDifference = DateDiff("d",objFile.DateLastModified,Now)

	If (dateDifference >= stockDay) Then
		objFile.Delete
		deleteOldFile = True
	Else
		deleteOldFile = False
	End If
End Function


