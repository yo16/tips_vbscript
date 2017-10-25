Option Explicit


Dim testStr1 'as String
Dim testStr2 'as String
testStr1 = "C:\900_Programming\VBScript\source\練習ソース\文字列操作"

testStr2 = popDirStr( testStr1, 1)

MsgBox(testStr2)




'Function popDirStr( pOldPath as String, pDelDepth as Integer ) As String
Function popDirStr( pOldPath , pDelDepth )
	' 戻り値の文字列
	Dim rtnStr 'As String
	rtnStr = pOldPath
	
	' 消したカウント
	Dim delCount 'As Integer
	delCount = 0
	
	
	Dim lastChar 'As String
	
	' １文字ずつ消すループ
	While ( delCount < pDelDepth )
		' 最後の文字を取得
		lastChar = Right( rtnStr, 1 )
		
		' なにはともあれ最後の文字をカット！
		rtnStr = Left( rtnStr, Len(rtnStr) - 1 )
		
		' 最後の文字が\かどうか判断する
		If ( lastChar = "\" ) Then
			' 最後の文字が「\」の場合、消したカウントをインクリメント
			delCount = delCount + 1
		End If
		
	Wend
	
	
	popDirStr = rtnStr
End Function







