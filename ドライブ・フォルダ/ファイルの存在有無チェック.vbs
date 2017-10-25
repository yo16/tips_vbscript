'******************************************
'[ファイルの存在有無チェック]
'
'２つのテキストファイルを見比べて
'重複しているもの、片方にしかないものを
'チェックする。
'比較結果はテキストファイルに出力する。
'
'比較結果の出力方法
'1 2
'* * sample1.txt	(両方のテキストファイルに存在)
'*   sample2.txt	()
'  * sample3.txt	()
'
'******************************************

Option Explicit

'比較するテキストファイル１
Dim textFile1
textFile2 = "file1.txt"
'比較するテキストファイル２
Dim textFile2
textFile2 = "file2.txt"
'比較結果を出力するファイル
Dim outputFile
outputFile = "cmpKekka.txt"


'ファイルシステムオブジェクト作成
Dim objFS
Set objFS = WScript.CreateObject("Scripting.FileSystemObject")
'出力用ファイルオブジェクト作成
Dim objWriteTS
Set objWriteTS = objFS.CreateTextFile(outputFile,True)
objWriteTS.WriteLine "1 2"


'ファイルを開くループで使う変数
Dim strLine, strTemp


'テキストファイル１を開く
Dim objTS
Set objTS = objFS.OpenTextFile(textFile1,1)

'テキストファイル１を１行ずつ読み
'テキストファイル２に存在するかチェックする
Do Until objTS.AtEndOfStream
	strTemp = objTS.ReadLine
	If Not(strTemp = "") Then
		If (strExistsInText(strTemp,textFile2)) Then
			'存在する場合
			objWriteTS.WriteLine "* * "&strTemp
		Else
			'存在しない場合
			objWriteTS.WriteLine "*   "&strTemp
		End If
	End If
Loop

'テキストファイル１を閉じる
objTS.Close


'テキストファイル２を開く
Set objTS = objFS.OpenTextFile(textFile2,1)

'テキストファイル２を１行ずつ読み
'テキストファイル１に存在するかチェックする
Do Until objTS.AtEndOfStream
	strTemp = objTS.ReadLine
	If Not(strTemp = "") Then
		If Not (strExistsInText(strTemp,textFile1)) Then
			'存在しない場合
			objWriteTS.WriteLine "  * "&strTemp
		End If
	End If
Loop

'テキストファイル２を閉じる
objTS.Close


MsgBox "終了～♪"




'関数：strExistsInText
'
'引数で渡すファイルに、
'指定された文字列が存在するか
'行単位で比較する。
'(「abc」と「abcd」ではFalseとなる。)
'存在する場合：Trueを返す
'存在しない場合：Falseを返す
Function strExistsInText(searchStr,searchFile)
	Dim objFS, objTS
	Dim strLine, strTemp

	Set objFS = WScript.CreateObject("Scripting.FileSystemObject")
	Set objTS = objFS.OpenTextFile(searchFile,1)
	strLine = ""

	Dim foundFlg
	foundFlg = False
	Do Until objTS.AtEndOfStream
		strTemp = objTS.ReadLine
		If (strTemp = searchStr) Then
			foundFlg = True
		End If
	Loop
	objTS.Close

	Return foundFlg
End Function
