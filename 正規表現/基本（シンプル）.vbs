Option Explicit

'���K�\����K�p�����镶�����ۑ�
Dim testStr
testStr = "youichirou.ikeda@excel.co.jp"

'�p�^�[�����쐬
Dim regPattern
regPattern = "^you[^\.]+"



'���K�\���I�u�W�F�N�g���쐬
Dim regEx
Set regEx = New RegExp

'���K�\���I�u�W�F�N�g�փ����o�ϐ��̐ݒ�
' - �p�^�[����ݒ�
regEx.Pattern = regPattern
' - ������S�̂���������悤�ɐݒ�iTrue�ɂ���ƁA�P�}�b�`����������Č�������j
regEx.Global = True

'Matches�I�u�W�F�N�g���쐬
dim Matches
Set Matches = regEx.Execute(testStr)


' ���@�@�s���|�C���g�łP�\��
MsgBox Matches.Item(0).Value

' ���@�A1���\��
Dim Match
For Each Match in Matches
	MsgBox Match.Value
Next
