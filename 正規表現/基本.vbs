Option Explicit

'正規表現を適用させる文字列を保存
Dim testStr
'testStr = "youichirou.ikeda@excel.co.jp"
testStr = InputBox("正規表現を適用させる文字列を入れてみてください。","にゃぁ")
If testStr = "" Then WScript.Quit

'パターンを作成
Dim regPattern
'regPattern = "a(.)\1"
regPattern = InputBox("正規表現のパターンを入れてみてください。","わんわん！")
If regPattern = "" Then WScript.Quit



'正規表現オブジェクトを作成
Dim regEx
Set regEx = New RegExp

'パターンを設定
regEx.Pattern = regPattern
'文字列全体を検索するように設定
regEx.Global = True

'Matchesオブジェクトを作成
Dim Matches
Set Matches = regEx.Execute(testStr)

'1つずつ表示
Dim Match
For Each Match in Matches
	MsgBox Match.Value
Next
