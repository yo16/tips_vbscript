Option Explicit



Dim YNmodori
YNmodori = MsgBox("���k�t���O�𗧂ĂĂ������ł����H",4,"���k�t���O�𗧂Ă�")
If (YNmodori <> 6) Then
	WScript.Quit
End If



msgbox compressFolder("a")


'********************************************
'�֐�:compressFolder
'����:folderName:���k����t�@�C����
'
'������������
'�t�@�C���̃v���p�e�B[���k(M)]���`�F�b�N����
'********************************************
Function compressFolder(folderName)
	'�t�@�C���̃v���p�e�B���擾����
	Dim objFS,objFolder
	Set objFS = WScript.CreateObject("Scripting.FileSystemObject")
	Set objFolder = objFS.GetFolder(folderName)
	Dim propertyValue
	propertyValue = objFolder.Attributes

msgbox "property"&propertyValue
	'���炩�Ɉ��k����Ă��Ȃ��ꍇ�͈��k�t���O�𗧂ĂĐ���I��
	If (propertyValue < 2048) Then
		objFolder.Attributes = propertyValue + 2048
		compressFolder = 0
		Exit Function
	End If

	'�v���p�e�B�̒l���Q�i���ɂ���
	Dim sho,amari,idx,propertyValue_2
	sho = propertyValue
	idx = 0
	Do Until (sho = 1)
		nishinNumber = nishinNumber + ( (sho mod 2) * (10^idx) )
		sho = sho \ 2
		idx = idx + 1
	Loop
	propertyValue_2 = propertyValue_2 + 10^idx

	'���k�̏�Ԃ������t���O���擾����
	Dim compressFlg
	compressFlg = Left(Right(propertyValue_2,11),1)

	'���k����Ă��Ȃ��ꍇ�͈��k�t���O�𗧂Ă�
	If (compressFlg = "0") Then
		objFolder.Attributes = propertyValue + 2048
	End If

	compressFolder = 0

End Function


