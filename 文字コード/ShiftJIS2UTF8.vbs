' ShiftJIS�̃t�@�C����UTF-8�֕ϊ�
' 2017/3/7 (c) yo16
' �����������Ƃ��߂����H

Option Explicit


Dim i
For i = 0 To WScript.Arguments.Count-1
	toUtf8 WScript.Arguments(i)
Next
msgbox "end"


Sub toUtf8(inFile)
	Dim outFile : outFile = inFile & "_utf8.txt"
	
	' ���̓t�@�C��
'	Dim objIn : Set objIn = CreateObject("ADODB.Stream")
'	objIn.Type = 2				' 1:�o�C�i�� | 2:�e�L�X�g
'	objIn.Charset = "iso-2022-jp"		' "UTF-8" | "iso-2022-jp" : ShiftJIS
'	objIn.Open
'	objIn.LoadFromFile inFile
	Dim objFs : Set objFs = CreateObject("Scripting.FileSystemObject")
	Dim objIn : Set objIn = objFs.OpenTextFile(inFile, 1)
	
	' �o�̓t�@�C��
	Dim objOut : Set objOut = CreateObject("ADODB.Stream")
	objOut.Type = 2
	objOut.Charset = "UTF-8"
	objOut.Open
	
	
	Dim line
'	Do Until objIn.EOS
	Do Until objIn.AtEndOfStream
'		line = objIn.ReadText(-2)	' -1:�S�s�ǂݍ��� | -2:�P�s�ǂݍ���
		line = objIn.ReadLine
'		msgbox line
		line = ExchangeHanKana2Wide(line)
		objOut.WriteText line, 1				' 0:������̂� | 1:������+���s
	Loop
	
	' �o�̓t�@�C���̕ۑ�
	objOut.SaveToFile outFile, 2		' 1:�w��t�@�C�����Ȃ���ΐV�K | 2:�t�@�C��������ꍇ�͏㏑��
	
	' �N���[�Y
	objIn.Close
	objOut.Close
End Sub

Function ExchangeHanKana2Wide(str)
	Dim aryHan, aryZen
	aryHan = Array( _
		"�","�","�","�","�","�", _
		"�","�","�","�","�","�","�","�","�","�", _
		"�","�","�","�","�", _
		"�","�","�","�","�", _
		"�","�","�","�","�", _
		"�","�","�","�","�", _
		"�","�","�","�","�", _
		"�","�","�","�","�", _
		"�","�","�","�","�", _
		"�","�","�","�","�", _
		"�","�","�","�","�", _
		"�","�")
	
	aryZen = Array( _
		"�B","�u","�v","�A","�E","��", _
		"�@","�B","�D","�F","�H","��","��","��","�b","�[", _
		"�A","�C","�E","�G","�I", _
		"�J","�L","�N","�P","�R", _
		"�T","�V","�X","�Z","�\", _
		"�^","�`","�c","�e","�g", _
		"�i","�j","�k","�l","�m", _
		"�n","�q","�t","�w","�z", _
		"�}","�~","��","��","��", _
		"��","��","��","��","��", _
		"��","��","��","��","��", _
		"�J","�K")

	' �S���̕����ɑ΂��Ēu�����Ăԁi�Ȃ񂾂��Ȃ��E�E�E�j
	Dim ub : ub = UBound(aryHan)
	Dim i
	For i=0 to ub
		str = Replace(str, aryHan(i), aryZen(i))
	Next
	
	ExchangeHanKana2Wide = str
End Function


