Option Explicit

'========================================'
'�֐�	indexOf
'����	�T���z��
'		�T���z��v�f
'�߂�l	���߂Ƀq�b�g�����z��ԍ�(0�x�[�X)
'		�q�b�g���Ȃ������ꍇ -1
'========================================'
Function indexOf(searchArray,searchString)
	Dim arrayValue
	Dim arrayIndex
	arrayIndex = 0
	For Each arrayValue In searchArray
		If (arrayValue = searchString) Then
			indexOf = arrayIndex
			Exit Function
		End If
		arrayIndex = arrayIndex + 1
	Next
	indexOf = -1
End Function
