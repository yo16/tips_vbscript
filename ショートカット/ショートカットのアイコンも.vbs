' アイコンも作る
' 2006/10/04 ikeda


Option Explicit

Dim objWshShell, objShortcut
Dim strDesktopPath
' WshShellオブジェクトを生成する
Set objWshShell = WScript.CreateObject("WScript.Shell")
' デスクトップのフォルダ名を取得する
strDesktopPath = objWshShell.SpecialFolders("Desktop")


Dim objFS
Set objFS = CreateObject("Scripting.FileSystemObject")
Dim objFile, objFileName, objFolder

Dim i, shortCutName
For i = 0 To WScript.Arguments.Count-1
	objFileName = WScript.Arguments(i)
	' この名前のものがファイルかフォルダか判断
	If (objFS.FolderExists(objFileName) = -1) Then
		' フォルダ
		Set objFolder = objFS.GetFolder(objFileName)
		shortCutName = objFolder.Name
	Else
		' ファイル
		Set objFile = objFS.GetFile(objFileName)
		shortCutName = objFile.Name
	End If

	' WshShortcutオブジェクトを生成する
	Set objShortcut = objWshShell.CreateShortcut(strDesktopPath & "\" & shortCutName & ".lnk")
	' ショートカットのターゲットファイルを指定する
	objShortcut.TargetPath = objFileName
	
	'-------------------------
	'アイコンを作る
	'-------------------------
'	objShortcut.IconLocation = "C:\900_Programming\VBScript\source\練習ソース\ショートカット\ICON0010.ICO, 0"
	objShortcut.IconLocation = "notepad.exe, 0"
	
	' ショートカットを作成する
	objShortcut.Save
Next



