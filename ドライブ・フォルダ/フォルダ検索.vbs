'--------------------------------------------------------
'フォルダを検索し、見つけたらファイルへ出力する
'--------------------------------------------------------
Option Explicit

'探す文字列
Dim findStr
findStr = "a"

Dim fileName, overWrite
fileName = "findFile.txt"
overWrite = True



Dim objFS, objTS
Set objFS = CreateObject("Scripting.FileSystemObject")
Set objTS = objFS.CreateTextFile(fileName, overWrite)


' ドロップしたフォルダの名前を出力
Dim objFolder, treeFolderName
Set objFolder = objFS.GetFolder(".")
treeFolderName = objFolder.Path


' フォルダ内をファイルに出力
printFolder treeFolderName


Sub printFolder(folderName)
	' 検索内容に一致するか確認
	Dim pos
	pos = Instr( objFS.GetFileName(folderName), findStr )
	If ( pos <> 0 ) Then
		' フォルダ名を出力
		objTS.WriteLine objFS.GetAbsolutePathName(folderName)
	End If
	
	' フォルダオブジェクトを取得
	Dim objFolder
	Set objFolder = objFS.GetFolder(folderName)
	
	' サブフォルダを取得
	Dim objSubFolders, objSubFolder
	Set objSubFolders = objFolder.SubFolders
	
	' サブフォルダの一覧を出力
	For Each objSubFolder In objSubFolders
		printFolder folderName&"\"&objSubFolder.Name
	Next
	
End Sub



