Option Explicit

Dim nishinSu
nishinSu = 1000

If (setNishinVal(nishinSu,3,1) = 0) Then
	MsgBox nishinSu,,"test"
End If


Function setNishinVal(byRef exNumber,exPos,exVal)
	On Error Resume Next

	Dim tmpNumber
	tmpNumber = CStr(CInt(exNumber))
	If Err Then
		WScript.Echo "Error!!!"
		setNishinVal = -1
		Exit Function
	End If

'	exNumber = Left(tmpNumber,(tmpNumber.length - (exPos + 1))) & Right(tmpNumber,exPos)
msgbox tmpNumber.length
length‚¶‚á‚È‚¢‚ç‚µ‚¢ 2001/02/13
	exNumber = Left(tmpNumber,(tmpNumber.length - 3))


'msgbox "test"&Left(tmpNumber,(tmpNumber.length - (exPos + 1)))
	setNishinVal = 0

End Function

