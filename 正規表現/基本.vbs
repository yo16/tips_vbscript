Option Explicit

'���K�\����K�p�����镶�����ۑ�
Dim testStr
'testStr = "youichirou.ikeda@excel.co.jp"
testStr = InputBox("���K�\����K�p�����镶��������Ă݂Ă��������B","�ɂ႟")
If testStr = "" Then WScript.Quit

'�p�^�[�����쐬
Dim regPattern
'regPattern = "a(.)\1"
regPattern = InputBox("���K�\���̃p�^�[�������Ă݂Ă��������B","�����I")
If regPattern = "" Then WScript.Quit



'���K�\���I�u�W�F�N�g���쐬
Dim regEx
Set regEx = New RegExp

'�p�^�[����ݒ�
regEx.Pattern = regPattern
'������S�̂���������悤�ɐݒ�
regEx.Global = True

'Matches�I�u�W�F�N�g���쐬
Dim Matches
Set Matches = regEx.Execute(testStr)

'1���\��
Dim Match
For Each Match in Matches
	MsgBox Match.Value
Next
