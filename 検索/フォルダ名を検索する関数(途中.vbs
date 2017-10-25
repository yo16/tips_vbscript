

'''''''''''''''''''''''''''''''''''''''''''''
'関数:searchFolder
'引数    P_FolderName:検索するファイル名
'        P_SearchFolderPath:検索先のフォルダ名
'               絶対パスでフォルダを指定(最後に\マークが必要)
'               サブフォルダも検索する(再帰)
'               ない場合は全Localドライブを検索する<<<<<未作成
'戻り値  ファイル名(絶対パス)
'               複数存在する場合でも、
'               はじめに見つけたファイルのみ返す
'''''''''''''''''''''''''''''''''''''''''''''
Function searchFolder(P_FolderName,P_SearchFolderPath)

	Dim objFS
	Set objFS = WScript.CreateObject("Scripting.FileSystemObject")

	'フォルダの存在チェック
	If objFS.FolderExists(P_SearchFolderPath & P_FolderName) Then
		searchFile = P_SearchFolderPath & P_FolderName
		Exit Function
	End If

	Dim objFolder
	Set objFolder = objFS.GetFolder(P_SearchFolderPath)

	Dim objSubFolders,objSubFolder
	Set objSubFolders = objFolder.SubFolders

	Dim rtnCode
	For Each objSubFolder In objSubFolders
		rtnCode = searchFile(P_FileName,P_FolderName&objSubFolder.Name&"\")
		If (rtnCode <> "") Then
			searchFile = rtnCode
			Exit Function
		End If
	Next

	searchFile = ""
End Function

