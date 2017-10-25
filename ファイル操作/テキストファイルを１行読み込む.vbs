
dim modori,rtnStr
modori = readOneLine("sample.txt",5,rtnStr)

msgbox rtnStr


''''''''''''''''''''''''''''''''''''''''''''''''
'�֐�:readOneLine
'����     p_fileName:�ǂݍ��ރt�@�C����
'         p_lineNumber:�ǂݍ��ލs�ԍ�
'         returnString:�ǂݍ��񂾕�����
'�߂�l   ����I��:0
'         �ُ�I��:-1
'
'���� ���� ����
'  �Ep_fileName��p_lineNumber�s�ڂ�ǂݍ���
'    �ǂݍ��񂾌��ʂ�Ԃ��֐�
'  �E�t�@�C�������݂��Ȃ��ꍇ�̓G���[
'  �E�t�@�C���̍s��>p_lineNumber �̏ꍇ�̓G���[
'2001/02/09 ikeda �쐬
'''''''''''''''''''''''''''''''''''''''''''''''
Function readOneLine(byRef p_fileName,p_lineNumber,returnString)
	On Error Resume Next

	'--  �������擾�ł��Ȃ��ꍇ�̃G���[����
	If ( (p_fileName = "") or (p_lineNumber = "") ) Then
		WScript.Echo "readOneLine:�������擾�ł��܂���ł����B" & Now
		readOneLine = -1
		Exit Function
	End If

	'--  �s�ԍ��������ȊO�������ꍇ�̃G���[����
	Dim tmpNumber
	tmpNumber = CInt(p_lineNumber)
	If Err Then
		WScript.Echo "readOneLine:�s�ԍ��͐������w�肵�Ă��������B" & Now
		readOneLine = -1
		Exit Function
	End If

	'--  �s�ԍ����}�C�i�X�̏ꍇ�̃G���[����
	If (CInt(p_lineNumber) <= 0) Then
		WScript.Echo "readOneLine:�w�肷��s�ԍ��͐��̐�������͂��Ă��������B" & Now
		readOneLine = -1
		Exit Function
	End If

	'--  �t�@�C�������݂��Ȃ��ꍇ�̃G���[����
	Dim objFS
	Set objFS = WScript.CreateObject("Scripting.FileSystemObject")
	If Not (objFS.FileExists(p_fileName)) Then
		WScript.Echo "readOneLine:�t�@�C��["&p_fileName&"]�����݂��܂���B" & Now
		readOneLine = -1
		Exit Function
	End If

	'--  �t�@�C�����J��
	Dim objTS
	Set objTS = objFS.OpenTextFile(p_fileName,1)
	If Err Then
		WScript.Echo "readOneLine:�t�@�C��["&p_fileNmae&"]���J�����Ƃ��ł��܂���ł����B" & Now
		readOneLine = -1
		Exit Function
	End If

	'--  �t�@�C����ǂ�
	Dim idx
	For idx = 1 to (p_lineNumber - 1)
		objTS.SkipLine
		If (objTS.AtEndOfStream = True) Then
			WScript.Echo "readOneLine:�w�肳�ꂽ�s�ԍ��̓t�@�C���̍s�������������ߓǂݍ��ނ��Ƃ��ł��܂���B" & Now
			readOneLine = -1
			Exit Function
		End If
		If Err Then
			WScript.Echo "readOneLine:�t�@�C��["&p_fileNmae&"]��ǂނ��Ƃ��ł��܂���ł����B" & Now
			readOneLine = -1
			Exit Function
		End If
	Next
	Dim tmpLine
	tmpLine = objTS.ReadLine
	objTS.Close
	If Err Then
		WScript.Echo "readOneLine:�t�@�C��["&p_fileNmae&"]��ǂނ��Ƃ��ł��܂���ł����B" & Now
		readOneLine = -1
		Exit Function
	End If

	'--  �߂�l��ݒ�
	returnString = tmpLine
	readOneLine = 0

End Function


