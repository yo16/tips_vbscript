' �֐��ɓn���Ă݂�e�X�g


' �ȒP�ɂł�������� 2007/04/24 



Option Explicit


' �z���`
Dim array1
array1 = Array("a","b","c")		'Array�֐����g�p

' �Ă�
test(array1)


' �֐���`
Sub test(pArray)
	' ���𐔂��Ă݂�
	msgbox UBound(pArray)
	
	' �o�͂��Ă݂�
	msgbox pArray(0) & "-" & pArray(1) & "-" & pArray(2)


End Sub

