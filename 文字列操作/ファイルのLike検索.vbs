Option Explicit
' ファイル名をlike検索する
' 2006/12/22 ikeda

Dim path1

path1 = "C:\900_Programming\VBScript\source\練習ソース\文字列操作\文字列*"


Dim objFS, objFolder, colFiles
Dim strFilesName, x
' FileSystemObject オブジェクトを生成する
Set objFS = WScript.CreateObject("Scripting.FileSystemObject")
' カレントフォルダのFolderオブジェクトを取得する
Set objFolder = objFS.GetFolder(".")
' カレントフォルダに含まれるすべてのファイルを取得する
Set colFiles = objFolder.Files
' 個々のファイル名を文字列に追加する
strFilesName = ""
For Each x in colFiles
	strFilesName = strFilesName & x.Name & vbCRLF
Next
' 結果を表示する
WScript.Echo strFilesName
