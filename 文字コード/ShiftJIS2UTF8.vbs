' ShiftJISのファイルをUTF-8へ変換
' 2017/3/7 (c) yo16
' 文字が長いとだめかも？

Option Explicit


Dim i
For i = 0 To WScript.Arguments.Count-1
	toUtf8 WScript.Arguments(i)
Next
msgbox "end"


Sub toUtf8(inFile)
	Dim outFile : outFile = inFile & "_utf8.txt"
	
	' 入力ファイル
'	Dim objIn : Set objIn = CreateObject("ADODB.Stream")
'	objIn.Type = 2				' 1:バイナリ | 2:テキスト
'	objIn.Charset = "iso-2022-jp"		' "UTF-8" | "iso-2022-jp" : ShiftJIS
'	objIn.Open
'	objIn.LoadFromFile inFile
	Dim objFs : Set objFs = CreateObject("Scripting.FileSystemObject")
	Dim objIn : Set objIn = objFs.OpenTextFile(inFile, 1)
	
	' 出力ファイル
	Dim objOut : Set objOut = CreateObject("ADODB.Stream")
	objOut.Type = 2
	objOut.Charset = "UTF-8"
	objOut.Open
	
	
	Dim line
'	Do Until objIn.EOS
	Do Until objIn.AtEndOfStream
'		line = objIn.ReadText(-2)	' -1:全行読み込み | -2:１行読み込み
		line = objIn.ReadLine
'		msgbox line
		line = ExchangeHanKana2Wide(line)
		objOut.WriteText line, 1				' 0:文字列のみ | 1:文字列+改行
	Loop
	
	' 出力ファイルの保存
	objOut.SaveToFile outFile, 2		' 1:指定ファイルがなければ新規 | 2:ファイルがある場合は上書き
	
	' クローズ
	objIn.Close
	objOut.Close
End Sub

Function ExchangeHanKana2Wide(str)
	Dim aryHan, aryZen
	aryHan = Array( _
		"｡","｢","｣","､","･","ｦ", _
		"ｧ","ｨ","ｩ","ｪ","ｫ","ｬ","ｭ","ｮ","ｯ","ｰ", _
		"ｱ","ｲ","ｳ","ｴ","ｵ", _
		"ｶ","ｷ","ｸ","ｹ","ｺ", _
		"ｻ","ｼ","ｽ","ｾ","ｿ", _
		"ﾀ","ﾁ","ﾂ","ﾃ","ﾄ", _
		"ﾅ","ﾆ","ﾇ","ﾈ","ﾉ", _
		"ﾊ","ﾋ","ﾌ","ﾍ","ﾎ", _
		"ﾏ","ﾐ","ﾑ","ﾒ","ﾓ", _
		"ﾔ","ﾕ","ﾖ","ﾗ","ﾘ", _
		"ﾙ","ﾚ","ﾛ","ﾜ","ﾝ", _
		"ﾞ","ﾟ")
	
	aryZen = Array( _
		"。","「","」","、","・","ヲ", _
		"ァ","ィ","ゥ","ェ","ォ","ャ","ュ","ョ","ッ","ー", _
		"ア","イ","ウ","エ","オ", _
		"カ","キ","ク","ケ","コ", _
		"サ","シ","ス","セ","ソ", _
		"タ","チ","ツ","テ","ト", _
		"ナ","ニ","ヌ","ネ","ノ", _
		"ハ","ヒ","フ","ヘ","ホ", _
		"マ","ミ","ム","メ","モ", _
		"ヤ","ユ","ヨ","ラ","リ", _
		"ル","レ","ロ","ワ","ン", _
		"゛","゜")

	' 全部の文字に対して置換を呼ぶ（なんだかなぁ・・・）
	Dim ub : ub = UBound(aryHan)
	Dim i
	For i=0 to ub
		str = Replace(str, aryHan(i), aryZen(i))
	Next
	
	ExchangeHanKana2Wide = str
End Function


