Option Explicit

'���K�\����K�p�����镶�����ۑ�
Dim testStr
testStr = "youichirou.ikeda@excel.co.jp"

'�p�^�[�����쐬
Dim regPattern
regPattern = "ou[^r]+"



'���K�\���I�u�W�F�N�g���쐬
Dim regEx
Set regEx = New RegExp

'���K�\���I�u�W�F�N�g�փ����o�ϐ��̐ݒ�
' - �p�^�[����ݒ�
regEx.Pattern = regPattern
' - ������S�̂���������悤�ɐݒ�
regEx.Global = True

'Matches�I�u�W�F�N�g���쐬
dim Matches
Set Matches = regEx.Execute(testStr)

'1���\��
Dim Match
Dim matchCount,i
matchCount = Matches.Count
MsgBox matchCount
For i=0 to (matchCount-1)
	Set Match = Matches.Item(i)
	MsgBox Match.Value
Next
