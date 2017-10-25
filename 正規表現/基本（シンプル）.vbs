Option Explicit

'正規表現を適用させる文字列を保存
Dim testStr
testStr = "youichirou.ikeda@excel.co.jp"

'パターンを作成
Dim regPattern
regPattern = "^you[^\.]+"



'正規表現オブジェクトを作成
Dim regEx
Set regEx = New RegExp

'正規表現オブジェクトへメンバ変数の設定
' - パターンを設定
regEx.Pattern = regPattern
' - 文字列全体を検索するように設定（Trueにすると、１つマッチした次から再検索する）
regEx.Global = True

'Matchesオブジェクトを作成
dim Matches
Set Matches = regEx.Execute(testStr)


' 方法①ピンポイントで１つ表示
MsgBox Matches.Item(0).Value

' 方法②1つずつ表示
Dim Match
For Each Match in Matches
	MsgBox Match.Value
Next
