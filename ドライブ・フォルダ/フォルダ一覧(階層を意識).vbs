'階層をたどって、
'ファイル・フォルダ一覧を
'テキストファイルに出力する
Option Explicit

Dim fileName, overWrite
fileName = "fileTree.txt"
overWrite = True
' ファイル出力フラグ
Dim fileFlag
fileFlag = False



Dim objFS, objTS
Set objFS = CreateObject("Scripting.FileSystemObject")
Set objTS = objFS.CreateTextFile(fileName, overWrite)


' ドロップしたフォルダの名前を出力
Dim objFolder, treeFolderName
Set objFolder = objFS.GetFolder(".")
treeFolderName = objFolder.Path


' フォルダ内をファイルに出力
printFolder treeFolderName, 0


Sub printFolder(folderName, floorNum)
	'フロア数(深さ)の分、タブを出力
	Dim i
	For i = 0 to (floorNum-1)
		objTS.Write vbTab
	Next

	' フォルダ名を出力
	objTS.WriteLine objFS.GetFileName(folderName)

	' サブフォルダを取得
	Dim objFolder
	Set objFolder = objFS.GetFolder(folderName)
	Dim objSubFolders, objSubFolder
	Set objSubFolders = objFolder.SubFolders

	' サブフォルダの一覧を出力
	For Each objSubFolder In objSubFolders
		printFolder folderName&"\"&objSubFolder.Name, floorNum+1
	Next

	If ( fileFlag ) Then
		' ファイルを出力
		Dim folFiles
		Set folFiles = objFolder.Files
		Dim objFiles
		For Each objFiles in folFiles
			For i = 0 to floorNum
				objTS.Write vbTab
			Next
			objTS.WriteLine objFiles.Name
		Next
	End If
	
End Sub



