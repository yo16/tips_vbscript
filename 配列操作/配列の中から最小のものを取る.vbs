Option Explicit


Dim tmpArray
tmpArray = Array("b","a","c")

msgbox minArrayIndex(tmpArray)



Function minArrayIndex(pArray)
	Dim arrayValue,pIndex
	pIndex = 0

	Dim minValue,minIndex
	minValue = pArray(0)
	minIndex = 0

	For Each arrayValue In pArray
		If (arrayValue < minValue) Then
			minIndex = pIndex
		End If
		pIndex = pIndex + 1
	Next

	minArrayIndex = minIndex
End Function

