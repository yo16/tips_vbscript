Option Explicit
Dim objFS, objFolder, colSubFolders
Dim strFoldersName, strFoldersPath, x
' FileSystemObject オブジェクトを生成する
Set objFS = WScript.CreateObject("Scripting.FileSystemObject")
' Folder オブジェクトを取得する
Set objFolder = objFS.GetFolder(".")
' サブフォルダの Folders コレクションを取得する
Set colSubFolders = objFolder.SubFolders
' すべてのサブフォルダ名をstrFoldersNameに入れる
strFoldersName = ""
strFoldersPath = ""
For Each x in colSubFolders
	' ファイル名
	strFoldersName = strFoldersName & x.Name & vbCRLF
	' パス
	strFoldersPath = strFoldersPath & x.Path & vbCrLf
Next
'WScript.Echo strFoldersName
WScript.Echo strFoldersPath
