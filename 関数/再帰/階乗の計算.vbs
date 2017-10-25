Option Explicit

MsgBox kaijou(4)


Function kaijou(numberX)
	If (numberX > 1) Then
		kaijou = numberX * kaijou(numberX -1)
	Else
		kaijou = 1
	End If
End Function


