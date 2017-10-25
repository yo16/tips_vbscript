Option Explicit

'変換する文字を入力
Dim exStr
exStr = InputBox("アスキーコードに変える文字列を"&VBCrLf&"入力してみてください！","入力してみ？")


'キャンセルor入力されていない場合
If (exStr = "") Then WScript.Quit


'文字数
'MsgBox "length = " & Len(exStr)

'１文字ずつ変換
Dim idx
Dim rtnStr
rtnStr = ""
For idx = 1 to Len(exStr)
	'１文字ずつ出力
	'MsgBox "文字" & idx & " = " & Mid(exStr,idx,1)

	'アスキーコードにして格納
	rtnStr = rtnStr & "Chr(" & Asc(Mid(exStr,idx,1)) & ")&"
Next

'最後の[&]を取る
rtnStr = Left(rtnStr,Len(rtnStr)-1)

'実行結果出力(MsgBox)
'MsgBox rtnStr

'実行結果出力(InputBox)
Dim modori
modori = InputBox("[" & exStr & "]を" & VBCrLf & "ASCIIコードに変えました！","結果発表～★",rtnStr)














