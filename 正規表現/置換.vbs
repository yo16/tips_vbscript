Option Explicit


Dim strTest
strTest = "aabbbbccaabbbbbbbbbbbcc"

Dim regExp
Set regExp = New RegExp
regExp.Pattern = "a+.*?c+"		' *の後に?が付いてると最短一致
regExp.Global = True			' 全部置換（ありなしで結果が異なる

Dim strReplace
strReplace = regExp.Replace( strTest, "X" )

MsgBox strReplace
