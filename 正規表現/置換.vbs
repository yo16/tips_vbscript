Option Explicit


Dim strTest
strTest = "aabbbbccaabbbbbbbbbbbcc"

Dim regExp
Set regExp = New RegExp
regExp.Pattern = "a+.*?c+"		' *�̌��?���t���Ă�ƍŒZ��v
regExp.Global = True			' �S���u���i����Ȃ��Ō��ʂ��قȂ�

Dim strReplace
strReplace = regExp.Replace( strTest, "X" )

MsgBox strReplace
