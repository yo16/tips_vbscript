Option Explicit


Dim array1
array1 = Array("a","b","c")
msgbox array1(2)
' ���̌��ɗv�f�������P�t�����������Ȃ��B�B


' �_��
'array1(3) = "d"
'msgbox array1(3)

' �_��
'array1 = Array(array1(0), array1(1), array1(2), "d")
'msgbox array1(3)

' �n�j!!!
Dim arrayTmp
arrayTmp = array1
array1 = null	'********�������d�v*******
array1 = Array(arrayTmp(0), arrayTmp(1), arrayTmp(2), "d")
msgbox array1(3)


