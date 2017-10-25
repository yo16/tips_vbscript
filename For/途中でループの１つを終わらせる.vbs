Option Explicit



Dim idx
Dim tmpNumber
tmpNumber = 0

For idx = 1 to 10
	tmpNumber = tmpNumber + 1
	If (idx >= 5) Then Exit For
Next

MsgBox "tmpNumber is "&tmpNumber



