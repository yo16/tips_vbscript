Option Explicit

Dim kaeruMoji
kaeruMoji = 111
If (maeZero(kaeruMoji,5) = 0) Then
	MsgBox kaeruMoji
End If



Function maeZero(byRef exStr,keta)
	Dim motoKeta
	motoKeta = Len(exStr)

	If (keta - motoKeta <= 0) Then
		WScript.Echo exStr&"��"&keta&"���ȏ�̂��߁A�����𒆒f���܂��B"
		maeZero = -1
		Exit Function
	End If

	Dim idx
	For idx = 1 to (keta - motoKeta)
		exStr = "0" & exStr
	Next

	maeZero = 0

End Function



