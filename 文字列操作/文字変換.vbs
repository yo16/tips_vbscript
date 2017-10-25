Option Explicit

Dim changeStr
changeStr = InputBox("Input Any Words!!")

Dim checkChar
checkChar = Asc(Mid(changeStr,1,1))

Dim ascChar
Dim idx,secondStr
secondStr = ""
For idx = 1 to len(changeStr)
	ascChar = Asc(Mid(changeStr,idx,1))
	If (CInt(ascChar) < 0) Then
		ascChar = ascChar * (-1)
	End If
	secondStr = secondStr & ascChar
Next









