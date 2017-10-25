Option Explicit
' ファイルをコピー２
' 存在し兄フォルダを指定した場合、勝手に作成してくれるか？

Dim objFS
Set objFS = WScript.CreateObject("Scripting.FileSystemObject")

'objFS.CopyFile "abc.txt",".\a\b\c\コピ～abc.txt"
' →エラー・・・

CreateFolder ".\a\b\c\コピ～abc.txt"
CreateFolder ".\a\b\c\d\e\"



' パスの下までフォルダを作成する
' 最後の文字が\の場合は、    最後のトークンをフォルダとみなし
'             \でない場合は、最後のトークンをファイルとみなす
Sub CreateFolder( strInputPath )
	Dim aryDir
	aryDir = Split(strInputPath, "\")
	Dim nLastIndex
	If( Right(strInputPath,1) = "\" ) Then
		' 最後はフォルダ
		nLastIndex = UBound(aryDir)
	Else
		' 最後はファイル
		nLastIndex = UBound(aryDir) - 1
	End If
	Dim strPath
	strPath = "."
	Dim i
	For i = 0 to nLastIndex
		strPath = strPath & "\" & aryDir(i)
		msgbox strPath
		If( objFS.FolderExists(strPath) = 0 )Then
			objFS.CreateFolder(strPath)
		End If
	Next
End Sub

