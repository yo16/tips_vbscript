Option Explicit

If (1 < 2) Then
	MsgBox "1 < 2"
End If

If (1 <= 2) Then
	MsgBox "1 <= 2"
End If

If (1 <> 2) Then
	MsgBox "1 <> 2"
End If

If (1 = 1) Then
	MsgBox "1 = 1"
End If

If (1 = 2) Then
	MsgBox "1 = 2"
Else
	MsgBox "1 <> 2"
End If


If (2 < 1) Then
	MsgBox "2 < 1"
ElseIf (2 > 1) Then
	MsgBox "2 > 1"
Else
	MsgBox "2 = 1"
End If


' ï°êîèåè
' ò_óùòaÅEò_óùêœ
If (1 = 1) and (2 = 2) Then
	MsgBox "(1 = 1) and (2 = 2)"
End If

If (1 = 1) or (2 <> 2) Then
	MsgBox "(1 = 1) or (2 <> 2)"
End If

' î€íË
If ( Not (1 = 1) ) Then
	MsgBox "Not (1 = 1)"
Else
	MsgBox "Not Not (1 = 1)"
End If

' Else if
If ( 1 = 2 ) Then
	MsgBox "1=2"
ElseIf ( 1 = 1 ) Then
	MsgBox "1=1"
Else
	MsgBox "else"
End If
