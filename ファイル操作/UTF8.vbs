' UTF-8のファイルを操作する
' このVBScripのソースは、shift-jisとする
' UTF-8にするにはwsfから呼び出す必要があるため、めんどくさい。（が、やればできる）

Option Explicit

Dim objFs
Set objFs = WScript.CreateObject("Scripting.FileSystemObject")

Dim testFileName
testFileName="utf-8test.txt"

' ------------------------------
' 新規作成
Dim objTs
Dim OverWrite
OverWrite = True
' unicodeファイルを作成するには、第３引数にTrueを設定する
Set objTs = objFs.CreateTextFile(testFileName, OverWrite, True)
objTs.WriteLine "日本語のテスト１２３"
objTs.WriteLine "日本語のテスト４５６"
objTs.Close
Set objTs = Nothing
' → これだと、utf-16になる！<< NG >>

Dim objFile
Set objFile = objFs.GetFile(testFileName)
Dim ForWriting
ForWriting = 2
Set objTs = objFile.OpenAsTextStream(ForWriting, -1)
objTs.WriteLine "日本語のテスト１２３"
objTs.WriteLine "日本語のテスト４５６"
objTs.Close
Set objTs = Nothing
' → これでも、utf-16になる！<< NG! >>

' これが解答
Dim outStream
Set outStream = CreateObject("ADODB.Stream")
outStream.type = 2	' 1:バイナリデータ | 2:テキストデータ
msgbox outStream.mode
outStream.mode = 0	' 1:読み取り | 2:書き込み | 3:読み取り/書き込み両方
' ↑なぜか3じゃないとできない。
'   または指定しなくてもできる。値は0。
outStream.charset = "UTF-8"
outStream.open
outStream.WriteText "日本語のテスト１２３", 0	' 第２引数：0:文字列を書き込む | 1:文字列＋改行文字を書き込む
outStream.WriteText "日本語のテスト４５６", 1
outStream.WriteText "日本語のテスト７８９", 1
' 保存
outStream.SaveToFile testFileName, 2	' 1:ファイルがない場合のみ作成 | 2:ある場合は上書き
outStream.close
Set outStream = Nothing

' ------------
' 追加書き込み
Dim addStream
Set addStream = CreateObject("ADODB.Stream")
addStream.type = 2
addStream.mode = 3	' 3:読み取り/書き込み両方
'   または指定しなくてもできる。値は0。
addStream.charset = "UTF-8"
addStream.open
addStream.LoadFromFile testFileName
addStream.Position = addStream.Size		' ポインタを終端へ
addStream.WriteText "追加です１２３", 1
addStream.SaveToFile testFileName, 2
addStream.close
Set addStream = Nothing



' ------------------------------
' 読み込み

' 引数がよくわからないけど、やっぱりこの関数じゃUTF-8は読めない
Set objTs = objFs.OpenTextFile(testFileName, 1,-1)
MsgBox objTs.ReadLine
MsgBox objTs.ReadLine
objTs.Close
Set objTs = Nothing

' これが解答
Dim inStream
Set inStream = CreateObject("ADODB.Stream")
inStream.type = 2
inStream.mode = 3
' ↑なぜか３じゃないとできない
'   または指定しなくてもできる。値は0。
inStream.charset = "UTF-8"
inStream.open
inStream.LoadFromFile testFileName
Do While inStream.EOS = False
	MsgBox inStream.ReadText(-2)	' -1:全部読み込む | -2:１行読み込む
Loop
inStream.close
Set inStream = Nothing

Set objFs = Nothing
