' 正規表現にマッチするかどうかだけ知りたい
' 2006/03/09


Dim regEx
Set regEx = New RegExp
regEx.Pattern = "abC"

If regEx.Test( "abc" ) Then
	msgbox "マッチ！"
Else
	msgbox "アンマッチ！"
End If

