' 問連123
' リクレポ123
' トラレポ123を全部検索できるようなマッチ文字列

Dim str1 : str1 = "問連123"
Dim str2 : str2 = "リクレポ456"
Dim str3 : str3 = "トラレポ789"



'正規表現オブジェクトを作成
Dim regEx
Set regEx = New RegExp

'パターンを設定
regEx.Pattern = "((問連)|(リクレポ)|(トラレポ))([0-9]+)"

'Matchesコレクションに入る
Dim Matches
Set Matches = regEx.Execute(str1)
MsgBox "str1.Count:" & Matches.Item(0).Value

Set Matches = regEx.Execute(str2)
MsgBox "str2.Count:" & Matches.Item(0).Value

Set Matches = regEx.Execute(str3)
MsgBox "str3.Count:" & Matches.Item(0).Value
MsgBox Matches.Item(0).SubMatches(4)
