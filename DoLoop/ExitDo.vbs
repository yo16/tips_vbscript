'Exit Do‚Á‚ÄŒ¾‚Á‚½uŠÔ‚Éƒ‹[ƒv‚©‚ço‚é
Option Explicit


Dim idx,str
idx = 0
str = ""

Do While idx<10
	If (idx = 5) Then
		Exit Do
	End If
	str = str & idx
	idx = idx + 1
Loop

msgbox str
