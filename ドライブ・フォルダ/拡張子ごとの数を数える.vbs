Option Explicit


Dim objFs
Set objFs = CreateObject("Scripting.FileSystemObject")
Dim objShell
Set objShell = CreateObject("WScript.Shell")

' カレントフォルダをこのスクリプトがあるフォルダにする
' (Drag&Dropしたときに変わってしまう対応)
objShell.CurrentDirectory = objFs.GetParentFolderName(WScript.ScriptFullName)

' 引数から、処理対象のフォルダを取得
If WScript.Arguments.Count = 0 Then
	WScript.Exit
End If
Dim targetDir
targetDir = WScript.Arguments(0)




Dim dicExt
Set dicExt = CreateObject("Scripting.Dictionary")

' 実行
CalcExtCount targetDir

Dim msg
msg = ""
Dim key
For Each key In dicExt
	msg = msg & key & ":" & dicExt.Item(key) & vbCrLf
Next
msgbox msg




Sub CalcExtCount( dirPath )
	Dim objDir
	Set objDir = objFs.GetFolder(dirPath)
	
	Dim objFile
	Dim ext
	For Each objFile In objDir.Files
		ext = Mid(objFile.Name, InStr(objFile.Name, ".")+1)
		If dicExt.Exists(ext) Then
			dicExt.Item(ext) = dicExt.Item(ext) + 1
		Else
			dicExt.Add ext, 1
		End If
	Next
	
	Dim objSubDir
	For Each objSubDir In objDir.SubFolders
		CalcExtCount objSubDir.Path
	Next
	
End Sub

