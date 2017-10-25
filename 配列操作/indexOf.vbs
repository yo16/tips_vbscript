Option Explicit

'========================================'
'関数	indexOf
'引数	探す配列
'		探す配列要素
'戻り値	初めにヒットした配列番号(0ベース)
'		ヒットしなかった場合 -1
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
