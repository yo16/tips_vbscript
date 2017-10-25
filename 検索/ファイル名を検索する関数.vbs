Option Explicit


Dim startTime
startTime = Time


Dim objFS
Dim strLine, strTemp

Set objFS = WScript.CreateObject("Scripting.FileSystemObject")

Dim objTSread,objTSwrite
'---  探すファイルが書かれているファイル
Set objTSread = objFS.OpenTextFile("存在チェック.dat",1)
'---  検索結果が書かれるファイル
Set objTSwrite = objFS.CreateTextFile("検索結果.txt",True)

strLine = ""
Dim searchFolderName
'---  検索するディレクトリ(サブフォルダも探す)
searchFolderName = "F:\Prsmhome\"

Do Until objTSread.AtEndOfStream
	strTemp = objTSread.ReadLine
	If Not(strTemp = "") Then
		objTSwrite.WriteLine strTemp&","&searchFile(strTemp,searchFolderName)
	End If
Loop

Dim endTime
endTime = Time
objTSwrite.WriteLine "startTime:" & startTime & "  endTime:" & endTime


objTSread.Close
objTSwrite.Close

'MsgBox "終わったよー。"


'''''''''''''''''''''''''''''''''''''''''''''
'関数:searchFile
'引数    P_FileName:検索するファイル名
'        P_FolderName:検索先のフォルダ名(最後に\マークが必要)
'               絶対パスでフォルダを指定
'               サブフォルダも検索する(再帰)
'               ない場合は全Localドライブを検索する<<<<<<<未作成
'戻り値  ファイル名(絶対パス)
'               複数存在する場合でも、
'               はじめに見つけたファイルのみ返す
'''''''''''''''''''''''''''''''''''''''''''''
Function searchFile(P_FileName,P_FolderName)

	Dim objFS
	Set objFS = WScript.CreateObject("Scripting.FileSystemObject")

	'ファイルの存在チェック
	If objFS.FileExists(P_FolderName & P_FileName) Then
		searchFile = P_FolderName & P_FileName
		Exit Function
	End If

	Dim objFolder
	Set objFolder = objFS.GetFolder(P_FolderName)

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

