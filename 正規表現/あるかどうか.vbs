' ���K�\���Ƀ}�b�`���邩�ǂ��������m�肽��
' 2006/03/09


Dim regEx
Set regEx = New RegExp
regEx.Pattern = "abC"

If regEx.Test( "abc" ) Then
	msgbox "�}�b�`�I"
Else
	msgbox "�A���}�b�`�I"
End If

