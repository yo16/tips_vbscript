' （）とか使ってみたり
' 2008/12/02 y.ikeda

Dim strTest
strTest = "abbbcccddddd"

Dim regEx
Set regEx = New RegExp
regEx.Pattern = "b+(c+)d+"

'Matchesオブジェクトを作成
Dim Matches
Set Matches = regEx.Execute(strTest)

'1つずつ表示
Dim Match, subMatch
For Each Match in Matches
	MsgBox Match.Value
	For Each subMatch In Match.SubMatches
		MsgBox subMatch
	Next
Next

