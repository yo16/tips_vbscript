Option Explicit

Dim searchMsg
searchMsg = InputBox("検索する文字列を入力してください。")
Dim replaceMsg
replaceMsg = InputBox("置換する文字列を入力してください。")


Dim objFS,objFolder,colFiles
Dim x
' FileSystemObject オブジェクトを生成する
Set objFS = WScript.CreateObject("Scripting.FileSystemObject")
' カレントフォルダのFolderオブジェクトを取得する
Set objFolder = objFS.GetFolder(".")
' カレントフォルダに含まれるすべてのファイルを取得する
Set colFiles = objFolder.Files


Dim debugStr


Dim objTS
For Each x in colFiles
	debugStr = ""
	Set objTS = x.OpenAsTextStream
	Do Until objTS.AtEndOfStream
		debugStr = debugStr & objTS.ReadLine & VBCrLf
	Loop
	msgbox debugStr
Next


