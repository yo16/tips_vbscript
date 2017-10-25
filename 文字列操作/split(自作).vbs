Option Explicit

' split�֐�(perl���ۂ�)

Dim motoStr
motoStr = "param=12345=abc"
'          123456789012345

Dim nameStr
Dim valueStr
Dim optionStr
nameStr = split( motoStr, "=", 1 )
msgbox nameStr

valueStr = split( motoStr, "=", 2 )
msgbox valueStr

optionStr = split( motoStr, "=", 3 )
msgbox optionStr


' str��sep�ŋ�؂���num�Ԗڂ̕������Ԃ�
' ��num��1����n�܂鐮��
Function split(str, sep, num)
	' �߂�l
	Dim returnStr
	
	' ���[�v�C���f�b�N�X
	Dim idx
	idx = 0
	' �����J�n�ʒu
	Dim startPos
	' �������ʒu
	Dim endPos
	endPos = 0
	' �������t���O
	Dim found
	
	While ( idx < num )
		' ����
		startPos = endPos+1
		endPos = Instr( startPos, str, sep )
		
		If ( endPos = 0 ) Then
'msgbox "not found"
			' �݂���Ȃ�����
			found = false
			' �܂������������Ȃ瑦�I��
			If ( idx+1 < num ) Then
				split = ""
			End If
		Else
'msgbox "found"
			' �݂����������݂͂����������猟��
			found = true
		End If
		
		' �C���N�������g
		idx = idx + 1
	Wend
	
	If ( found ) Then
		' �Ō�A�݂��Ă���AstartPos�`endPos
		returnStr = Mid( str, startPos, endPos-startPos )
	Else
		' �Ōオ�݂���Ȃ�������AstartPos�`������̍Ō�
'		returnStr = Mid( str, startPos, Len(str)-startPos+1 )
		returnStr = Mid( str, startPos )
	End If
	
	split = returnStr
	
End Function
