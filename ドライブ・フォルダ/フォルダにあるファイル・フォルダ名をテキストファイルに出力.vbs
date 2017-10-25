Option Explicit

Dim fileName
'********出力ファイル名********
fileName = "fileNames.txt"
Dim overWrite
overWrite = True


Dim objFS, objTS
Set objFS = WScript.CreateObject("Scripting.FileSystemObject")
Set objTS = objFS.CreateTextFile(fileName,overWrite)

Dim objFolder
Set objFolder = objFS.GetFolder(".")

'ファイル一覧を作成
objTS.WriteLine "** ファイル一覧 **"
Dim colFiles
Set colFiles = objFolder.Files
Dim x
For Each x in colFiles
	If ((x.Name <> fileName) and (x.Name <> WScript.ScriptName)) Then
		objTS.WriteLine x.Name
	End If
Next

'フォルダ一覧を作成
objTS.WriteLine "** フォルダ一覧 **"
Dim colSubFolders
Set colSubFolders = objFolder.SubFolders
For Each x in colSubFolders
	objTS.WriteLine x.Name
Next


objTS.Close

