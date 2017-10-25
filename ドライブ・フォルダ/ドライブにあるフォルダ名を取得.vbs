Option Explicit


Dim objFS, objFolder, colSubFolders
Dim strFoldersName, x

' FileSystemObject オブジェクトを生成する
Set objFS = WScript.CreateObject("Scripting.FileSystemObject")

' Folder オブジェクトを取得する
Set objFolder = objFS.GetFolder(".")

' サブフォルダの Folders コレクションを取得する
Set colSubFolders = objFolder.SubFolders

' すべてのサブフォルダ名をstrFoldersNameに入れる
strFoldersName = ""
For Each x in colSubFolders
	strFoldersName = strFoldersName & x.Name & vbCRLF
Next

WScript.Echo strFoldersName



