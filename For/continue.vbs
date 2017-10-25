Option Explicit

' for continue‚Ì‚æ‚¤‚È‚à‚Ì‚Í
' VBS‚É‚Í‚È‚¢


Dim idx
Dim tmpNumber
tmpNumber = 0

For idx = 1 to 10 : Do
	If (idx > 5) Then Exit Do
	tmpNumber = tmpNumber + 1
Loop Until 1: Next

MsgBox "tmpNumber is "&tmpNumber


' ‰º‹L‚ÌFor‚ÆDo‚ð‚Ps‚Å‘‚¢‚½‚¾‚¯
'	For idx = 1 to 10
'		Do
'			If (idx > 5) Then Exit Do
'			tmpNumber = tmpNumber + 1
'		Loop Until 1
'	Next


