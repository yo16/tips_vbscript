'******************************************
'[�t�@�C���̑��ݗL���`�F�b�N]
'
'�Q�̃e�L�X�g�t�@�C��������ׂ�
'�d�����Ă�����́A�Е��ɂ����Ȃ����̂�
'�`�F�b�N����B
'��r���ʂ̓e�L�X�g�t�@�C���ɏo�͂���B
'
'��r���ʂ̏o�͕��@
'1 2
'* * sample1.txt	(�����̃e�L�X�g�t�@�C���ɑ���)
'*   sample2.txt	()
'  * sample3.txt	()
'
'******************************************

Option Explicit

'��r����e�L�X�g�t�@�C���P
Dim textFile1
textFile2 = "file1.txt"
'��r����e�L�X�g�t�@�C���Q
Dim textFile2
textFile2 = "file2.txt"
'��r���ʂ��o�͂���t�@�C��
Dim outputFile
outputFile = "cmpKekka.txt"


'�t�@�C���V�X�e���I�u�W�F�N�g�쐬
Dim objFS
Set objFS = WScript.CreateObject("Scripting.FileSystemObject")
'�o�͗p�t�@�C���I�u�W�F�N�g�쐬
Dim objWriteTS
Set objWriteTS = objFS.CreateTextFile(outputFile,True)
objWriteTS.WriteLine "1 2"


'�t�@�C�����J�����[�v�Ŏg���ϐ�
Dim strLine, strTemp


'�e�L�X�g�t�@�C���P���J��
Dim objTS
Set objTS = objFS.OpenTextFile(textFile1,1)

'�e�L�X�g�t�@�C���P���P�s���ǂ�
'�e�L�X�g�t�@�C���Q�ɑ��݂��邩�`�F�b�N����
Do Until objTS.AtEndOfStream
	strTemp = objTS.ReadLine
	If Not(strTemp = "") Then
		If (strExistsInText(strTemp,textFile2)) Then
			'���݂���ꍇ
			objWriteTS.WriteLine "* * "&strTemp
		Else
			'���݂��Ȃ��ꍇ
			objWriteTS.WriteLine "*   "&strTemp
		End If
	End If
Loop

'�e�L�X�g�t�@�C���P�����
objTS.Close


'�e�L�X�g�t�@�C���Q���J��
Set objTS = objFS.OpenTextFile(textFile2,1)

'�e�L�X�g�t�@�C���Q���P�s���ǂ�
'�e�L�X�g�t�@�C���P�ɑ��݂��邩�`�F�b�N����
Do Until objTS.AtEndOfStream
	strTemp = objTS.ReadLine
	If Not(strTemp = "") Then
		If Not (strExistsInText(strTemp,textFile1)) Then
			'���݂��Ȃ��ꍇ
			objWriteTS.WriteLine "  * "&strTemp
		End If
	End If
Loop

'�e�L�X�g�t�@�C���Q�����
objTS.Close


MsgBox "�I���`��"




'�֐��FstrExistsInText
'
'�����œn���t�@�C���ɁA
'�w�肳�ꂽ�����񂪑��݂��邩
'�s�P�ʂŔ�r����B
'(�uabc�v�Ɓuabcd�v�ł�False�ƂȂ�B)
'���݂���ꍇ�FTrue��Ԃ�
'���݂��Ȃ��ꍇ�FFalse��Ԃ�
Function strExistsInText(searchStr,searchFile)
	Dim objFS, objTS
	Dim strLine, strTemp

	Set objFS = WScript.CreateObject("Scripting.FileSystemObject")
	Set objTS = objFS.OpenTextFile(searchFile,1)
	strLine = ""

	Dim foundFlg
	foundFlg = False
	Do Until objTS.AtEndOfStream
		strTemp = objTS.ReadLine
		If (strTemp = searchStr) Then
			foundFlg = True
		End If
	Loop
	objTS.Close

	Return foundFlg
End Function
