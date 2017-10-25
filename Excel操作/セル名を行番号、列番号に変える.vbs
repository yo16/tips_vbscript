Option Explicit

Dim strCellName
Dim nRow, nCol
Dim nRtn

strCellName = "aa3"
nRtn = GetXlsRowCol(strCellName, nRow, nCol)
MsgBox nRow & "-" & nCol


' エクセルのセル名を行番号、列番号に変える関数
' 戻り値	0:正常終了
' 2007/10/12
Function GetXlsRowCol( strXlsCellName, ByRef nXlsRow, ByRef nXlsCol )
	nXlsRow = 0
	nXlsCol = 0
	
	' 正規表現オブジェクトを設定
	Dim regEx
	Set regEx = New RegExp
	regEx.Pattern = "^[A-Za-z]+[0-9]+$"
	
	' 検索
	If ( Not regEx.Test( strXlsCellName ) ) Then
		' アンマッチ
		Msgbox "フォーマットが違ってそうです" & vbCrLf & strXlsCellName
		GetXlsRowCol = -1
		Exit Function
	End If
	
	' 大文字化
	strXlsCellName = UCase( strXlsCellName )
	
	' アルファベットと数字を分けるために再検索
	Dim nSepPos
	regEx.Pattern = "[0-9]+"
	Dim regMatches, regMatch
	Set regMatches = regEx.Execute( strXlsCellName )
	For Each regMatch in regMatches
		nSepPos = regMatch.FirstIndex
		Exit For
	Next
	
	' 分ける
	Dim strAlphabetPart, strNumberPart
	strAlphabetPart = Left( strXlsCellName, nSepPos )
	strNumberPart = Mid( strXlsCellName, nSepPos+1 )
	
	' アルファベットパートを数値に変換する（26進数→10進数）
	' 下の位から１文字ずつ変換していく
	Dim nPos, i
	Dim cAZ, strAbc
	strAbc = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
	For i = 1 to Len( strAlphabetPart )
		' １文字取得
		cAZ = Mid( strAlphabetPart, Len( strAlphabetPart )-i+1, 1 )
		' 数値化
		nPos = Instr( strAbc, cAZ )
		
		' 戻り値（列）へ設定
		nXlsCol = nXlsCol + nPos * ( Len(strAbc) ^ (i-1) )
	Next
	' 戻り値（行）へ設定
	nXlsRow = Int( strNumberPart )
	
	GetXlsRowCol = 0
End Function
