' �z����֐��֓n��

Option Explicit


' �z���`
Dim array1
array1 = Array("a","b","c")		'Array�֐����g�p

' �Ă�
test array1

test2 array1
msgbox array1(1)
' �� X
' �ύX����Ă���


' �֐���`
Sub test(pArray)
	' ���𐔂��Ă݂�
	msgbox UBound(pArray), vbOkOnly, "UBound(array)"
	' �� 2
	
	' �o�͂��Ă݂�
	msgbox pArray(0) & "-" & pArray(1) & "-" & pArray(2), vbOkOnly, "elements"
	' �� a-b-c
End Sub

' ByRef�Ŏ󂯎���ĕύX
Sub test2(ByRef pArray)
	pArray(1) = "X"
End Sub
