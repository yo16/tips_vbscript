's099.vbs

Option Explicit
Dim objFS, objFolder, colFiles
' FileSystemObject オブジェクトを生成する
Set objFS = WScript.CreateObject("Scripting.FileSystemObject")
' Folder オブジェクトを取得する
Set objFolder = objFS.GetFolder(".")
' Files オブジェクトを取得する
Set colFiles = objFolder.Files
' ファイルの個数を表示する
WScript.Echo colFiles.Count
