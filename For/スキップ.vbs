' ���[�v���X�L�b�v����
' C�����continue
' VBS�ɂ͑��݂��Ȃ��̂ŁA�ǂ����邩�H
' If�́A��΂�������������������ƁA�ǂ�ǂ�[���Ȃ邽��NG
Dim str
str = ""

Dim i
For i=0 to 10
Do
	If (i mod 2 = 0 ) Then Exit Do
	If (i mod 3 = 0 ) Then Exit Do
	str = str & "/" & i
Loop Until 1
Next

MsgBox str
' /1/5/7
' �O�`�P�O�ŁA�Q�̔{���ƂR�̔{���ȊO
