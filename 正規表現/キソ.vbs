' キソ

Dim str : str = "abc123xxx456def789"



'正規表現オブジェクトを作成
Dim regEx
Set regEx = New RegExp

'パターンを設定
regEx.Pattern = "([0-9])[a-z]"
'文字列全体を検索するように設定
regEx.Global = True

'Matchesコレクションに入る
Dim Matches
Set Matches = regEx.Execute(str)

'1つずつ表示
Dim Match, subMatch
For Each Match in Matches
	' マッチした全体は、Match.Valueに入っている
	MsgBox Match.Value
	
	' ()内の文字列を取りたい場合は、SubMatchesを使う
	For Each subMatch in Match.SubMatches
		MsgBox subMatch
		' .Valueでないので注意
	Next
	
Next

' マッチ数
MsgBox "マッチ数：" & Matches.Count

' 最初の１つしか使用しなければ、Itemも使える
MsgBox Matches.Item(0).Value
