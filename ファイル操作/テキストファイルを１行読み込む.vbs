
dim modori,rtnStr
modori = readOneLine("sample.txt",5,rtnStr)

msgbox rtnStr


''''''''''''''''''''''''''''''''''''''''''''''''
'関数:readOneLine
'引数     p_fileName:読み込むファイル名
'         p_lineNumber:読み込む行番号
'         returnString:読み込んだ文字列
'戻り値   正常終了:0
'         異常終了:-1
'
'＊＊ 説明 ＊＊
'  ・p_fileNameのp_lineNumber行目を読み込み
'    読み込んだ結果を返す関数
'  ・ファイルが存在しない場合はエラー
'  ・ファイルの行数>p_lineNumber の場合はエラー
'2001/02/09 ikeda 作成
'''''''''''''''''''''''''''''''''''''''''''''''
Function readOneLine(byRef p_fileName,p_lineNumber,returnString)
	On Error Resume Next

	'--  引数を取得できない場合のエラー処理
	If ( (p_fileName = "") or (p_lineNumber = "") ) Then
		WScript.Echo "readOneLine:引数を取得できませんでした。" & Now
		readOneLine = -1
		Exit Function
	End If

	'--  行番号が整数以外だった場合のエラー処理
	Dim tmpNumber
	tmpNumber = CInt(p_lineNumber)
	If Err Then
		WScript.Echo "readOneLine:行番号は整数を指定してください。" & Now
		readOneLine = -1
		Exit Function
	End If

	'--  行番号がマイナスの場合のエラー処理
	If (CInt(p_lineNumber) <= 0) Then
		WScript.Echo "readOneLine:指定する行番号は正の整数を入力してください。" & Now
		readOneLine = -1
		Exit Function
	End If

	'--  ファイルが存在しない場合のエラー処理
	Dim objFS
	Set objFS = WScript.CreateObject("Scripting.FileSystemObject")
	If Not (objFS.FileExists(p_fileName)) Then
		WScript.Echo "readOneLine:ファイル["&p_fileName&"]が存在しません。" & Now
		readOneLine = -1
		Exit Function
	End If

	'--  ファイルを開く
	Dim objTS
	Set objTS = objFS.OpenTextFile(p_fileName,1)
	If Err Then
		WScript.Echo "readOneLine:ファイル["&p_fileNmae&"]を開くことができませんでした。" & Now
		readOneLine = -1
		Exit Function
	End If

	'--  ファイルを読む
	Dim idx
	For idx = 1 to (p_lineNumber - 1)
		objTS.SkipLine
		If (objTS.AtEndOfStream = True) Then
			WScript.Echo "readOneLine:指定された行番号はファイルの行数よりも多いため読み込むことができません。" & Now
			readOneLine = -1
			Exit Function
		End If
		If Err Then
			WScript.Echo "readOneLine:ファイル["&p_fileNmae&"]を読むことができませんでした。" & Now
			readOneLine = -1
			Exit Function
		End If
	Next
	Dim tmpLine
	tmpLine = objTS.ReadLine
	objTS.Close
	If Err Then
		WScript.Echo "readOneLine:ファイル["&p_fileNmae&"]を読むことができませんでした。" & Now
		readOneLine = -1
		Exit Function
	End If

	'--  戻り値を設定
	returnString = tmpLine
	readOneLine = 0

End Function


