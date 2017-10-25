Option Explicit

Dim regExp
Set regExp = new RegExp
regExp.Pattern = "abc([0-9]+)?"



Dim str1
str1 = "abc1"


' 正規表現でマッチを実行
Dim matches
Set matches = regExp.Execute( str1 )


msgbox "見つけた数:" & matches.Count

Dim m
Dim sm
For Each m in matches
	msgbox "Match:" & m
	msgbox "submatchの数:" & m.SubMatches.Count
	For Each sm in m.SubMatches
		msgbox "SubMatch:" & sm
	Next
Next


' ループでまわさない方法
msgbox "まわさない："&matches(0).SubMatches(0)
