Option Explicit

Dim objDict
Set objDict = CreateObject("Scripting.Dictionary")

objDict.CompareMode = vbTextCompare
'Addメソッド
' 第一引数がKey、第二引数がItem
objDict.Add "1","あいうえお"
objDict.Add "2","かきくけこ"
objDict.Add "3","さしすせそ"
objDict.Add "4","たちつてと"

'ItemsメソッドとKeysメソッドを使ってみる
Dim strItems,strKeys
strItems = objDict.Items
strKeys = objDict.Keys

Dim idx
'Countプロパティを使ってみる
For idx = 0 To objDict.Count - 1
	MsgBox "キー "&strKeys(idx)&" に対応するデータは"&strItems(idx)&"です。"
Next

