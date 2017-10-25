Option Explicit
Dim objFS, objFolder, colFiles
Dim strFilesName, x
' FileSystemObject オブジェクトを生成する
Set objFS = WScript.CreateObject("Scripting.FileSystemObject")
' カレントフォルダのFolderオブジェクトを取得する
Set objFolder = objFS.GetFolder(".")
' カレントフォルダに含まれるすべてのファイルを取得する
Set colFiles = objFolder.Files
' 数を表示
MsgBox colFiles.count
