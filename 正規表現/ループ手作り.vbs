Option Explicit

'正規表現を適用させる文字列を保存
Dim testStr
testStr = "youichirou.ikeda@excel.co.jp"

'パターンを作成
Dim regPattern
regPattern = "ou[^r]+"



'正規表現オブジェクトを作成
Dim regEx
Set regEx = New RegExp

'正規表現オブジェクトへメンバ変数の設定
' - パターンを設定
regEx.Pattern = regPattern
' - 文字列全体を検索するように設定
regEx.Global = True

'Matchesオブジェクトを作成
dim Matches
Set Matches = regEx.Execute(testStr)

'1つずつ表示
Dim Match
Dim matchCount,i
matchCount = Matches.Count
MsgBox matchCount
For i=0 to (matchCount-1)
	Set Match = Matches.Item(i)
	MsgBox Match.Value
Next
