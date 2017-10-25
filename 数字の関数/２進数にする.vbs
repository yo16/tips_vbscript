Option Explicit
On Error Resume Next


Dim exNumber
Dim exNumberLength,idx,lengthMsg
exNumber = InputBox("”š‚ğ“ü‚ê‚Ä‚­‚¾‚³‚¢I")
If (toNishinSu(exNumber) = 0) Then
	exNumberLength = Len(exNumber)
	lengthMsg = ""
	For idx = 1 to exNumberLength
		lengthMsg = idx & lengthMsg
	Next
	msgBox lengthMsg & vbCrLf & exNumber
End If




Function toNishinSu(byRef changeNumber)
	changeNumber = CInt(changeNumber)
	If Err Then
		WScript.Echo "®”‚ª“ü—Í‚³‚ê‚Ä‚¢‚Ü‚¹‚ñI"
		toNishinSu = -1
		Exit Function
	End If

	Dim sho,amari,idx
	Dim nishinNumber
	sho = changeNumber
	idx = 0
	nishinNumber = 0
	Do Until (sho = 1)
		nishinNumber = nishinNumber + ( (sho mod 2) * (10^idx) )
		sho = sho \ 2
		idx = idx + 1
	Loop
	nishinNumber = nishinNumber + 10^idx

	changeNumber = nishinNumber
	toNishinSu = 0
End Function

