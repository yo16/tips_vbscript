Option Explicit
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
'	strFilesName = strFilesName & x.Name & vbCRLF
	strFilesName = strFilesName & x.Path & vbCRLF
Next
' 結果を表示する
WScript.Echo strFilesName
