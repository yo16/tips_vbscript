Dim OmiValue, Response, MsgStr
Randomize   ' 乱数ジェネレータを初期化。
OmiValue = Int((6 * Rnd) + 1)   ' 1 ～ 6 のランダムな値を生成。


If ( OmiValue = 1 ) Then
	MsgStr = "★大吉★"
ElseIf ( OmiValue = 2 ) Then
	MsgStr = "★中吉★"
ElseIf ( OmiValue = 3 ) Then
	MsgStr = "★小吉★"
ElseIf ( OmiValue = 4 ) Then
	MsgStr = "★末吉★"
ElseIf ( OmiValue = 5 ) Then
	MsgStr = "★凶★"
Else
	MsgStr = "★大凶★"
End If

MsgBox MsgStr, vbYes, "今日の運勢"



