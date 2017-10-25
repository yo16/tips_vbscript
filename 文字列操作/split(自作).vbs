Option Explicit

' split関数(perlっぽく)

Dim motoStr
motoStr = "param=12345=abc"
'          123456789012345

Dim nameStr
Dim valueStr
Dim optionStr
nameStr = split( motoStr, "=", 1 )
msgbox nameStr

valueStr = split( motoStr, "=", 2 )
msgbox valueStr

optionStr = split( motoStr, "=", 3 )
msgbox optionStr


' strをsepで区切ってnum番目の文字列を返す
' ※numは1から始まる整数
Function split(str, sep, num)
	' 戻り値
	Dim returnStr
	
	' ループインデックス
	Dim idx
	idx = 0
	' 検索開始位置
	Dim startPos
	' 見つけた位置
	Dim endPos
	endPos = 0
	' 見つけたフラグ
	Dim found
	
	While ( idx < num )
		' 検索
		startPos = endPos+1
		endPos = Instr( startPos, str, sep )
		
		If ( endPos = 0 ) Then
'msgbox "not found"
			' みつからなかった
			found = false
			' まだ検索が続くなら即終了
			If ( idx+1 < num ) Then
				split = ""
			End If
		Else
'msgbox "found"
			' みつかった→次はみつかった次から検索
			found = true
		End If
		
		' インクリメント
		idx = idx + 1
	Wend
	
	If ( found ) Then
		' 最後、みつけてたら、startPos～endPos
		returnStr = Mid( str, startPos, endPos-startPos )
	Else
		' 最後がみつからなかったら、startPos～文字列の最後
'		returnStr = Mid( str, startPos, Len(str)-startPos+1 )
		returnStr = Mid( str, startPos )
	End If
	
	split = returnStr
	
End Function
