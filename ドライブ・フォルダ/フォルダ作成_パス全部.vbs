' 指定したフォルダパスの、途中がなくても全部作る
Option Explicit



Dim objFS
Set objFS = WScript.CreateObject("Scripting.FileSystemObject")

Dim strPath
strPath = ".\aa\bb\cccc"

Call subCreDir(strPath)


Sub subCreDir(path)
	If objFS.FolderExists(path) Then
		exit sub
	End If
	
	' 後ろから\まで抜き出す(\含まず)
	Dim nSepPos
	nSepPos = InStrRev(path, "\")
	Dim strDirName
	strDirName = Mid(path, nSepPos+1)
	'msgbox strDirName
	
	' フォルダ上位フォルダの存在チェック
	If Not objFS.FolderExists(Left(path,nSepPos-1)) Then
		' 存在しない場合、上位フォルダを作成する
		Call subCreDir( Left(path,nSepPos-1) )
	End If
	
	' カレントのフォルダを作成する
	objFS.CreateFolder(path)
End Sub

