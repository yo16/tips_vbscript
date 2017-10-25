Option Explicit


' パターンは\はエスケープだけど、検索対象の文字列はエスケープでない
Dim regExp
Set regExp = new RegExp
regExp.Pattern = "a\\bc([0-9]+)?"



Dim str1
str1 = "a\bc1"


' 正規表現でマッチを実行
Dim matches
Set matches = regExp.Execute( str1 )


msgbox "見つけた数:" & matches.Count


