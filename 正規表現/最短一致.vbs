Option Explicit



Dim strTest
strTest = "aabbccaabbcc"

Dim regExp
Set regExp = New RegExp
'regExp.Pattern = "a+.*c+"
regExp.Pattern = "a+.*?c+"		' *�̌��?���t���Ă�ƍŒZ��v

Dim matches
Set matches = regExp.Execute( strTest )


Dim Match
For Each Match in Matches
	MsgBox Match.Value
Next
