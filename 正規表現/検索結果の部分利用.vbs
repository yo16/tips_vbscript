Option Explicit

Dim regExp
Set regExp = new RegExp
regExp.Pattern = "abc([0-9]+)?"



Dim str1
str1 = "abc1"


' ���K�\���Ń}�b�`�����s
Dim matches
Set matches = regExp.Execute( str1 )


msgbox "��������:" & matches.Count

Dim m
Dim sm
For Each m in matches
	msgbox "Match:" & m
	msgbox "submatch�̐�:" & m.SubMatches.Count
	For Each sm in m.SubMatches
		msgbox "SubMatch:" & sm
	Next
Next


' ���[�v�ł܂킳�Ȃ����@
msgbox "�܂킳�Ȃ��F"&matches(0).SubMatches(0)
