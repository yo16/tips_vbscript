Option Explicit


' �p�^�[����\�̓G�X�P�[�v�����ǁA�����Ώۂ̕�����̓G�X�P�[�v�łȂ�
Dim regExp
Set regExp = new RegExp
regExp.Pattern = "a\\bc([0-9]+)?"



Dim str1
str1 = "a\bc1"


' ���K�\���Ń}�b�`�����s
Dim matches
Set matches = regExp.Execute( str1 )


msgbox "��������:" & matches.Count


