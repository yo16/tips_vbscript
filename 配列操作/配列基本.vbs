Option Explicit

'�����ݒ�
Dim array0
array0 = Array()
' �I���̂Ƃ�
Set array0 = Nothing



'�g�p���@�P
Dim array1
array1 = Array("a","b","c")		'Array�֐����g�p

'�z��̓[���x�[�X��()���g��
msgbox array1(0) & "-" & array1(1) & "-" & array1(2)



'�g�p���@�Q
Dim array2(2)		'��-1��錾�i�[���̕��j
array2(0) = "A"
array2(1) = "B"
array2(2) = "C"

'�z��̓[���x�[�X��()���g��
msgbox array2(0) & "-" & array2(1) & "-" & array2(2)

' �v�f��
MsgBox "UBound(array2)=" & UBound(array2)

'�_���ȗ�P
'Dim array3			'�C�����͔z��(���ۂ͂O����)
'array3(0) = "a"		'�z��ɓ���Ă݂�

' ���߂ȗ�Q
'Dim array4
'array4 = Array(3)		' �v�f���ł��ĂȂ��B�ǂ�������Ԃ��s���B

' ���߂ȗ�R
Dim d1
d1 = 3
'Dim array5(d1)		' �v�f���̒�`�ɕϐ����g���Ȃ�
Dim array5
ReDim array5(d1)	' �������@�F��������^�Ȃ��Œ�`������AReDim�Ŕz��ɂ���
array5(0) = "x"


