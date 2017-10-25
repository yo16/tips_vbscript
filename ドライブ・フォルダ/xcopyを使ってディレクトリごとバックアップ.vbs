Option Explicit


'Dim YNmodori
'YNmodori = MsgBox("バックアップを取ってもいいですか？",4,"フォルダごとバックアップ")
'If (YNmodori <> 6) Then
'	WScript.Quit
'End If


Dim motoBKFolderName
motoBKFolderName = "MartBrowserSourceBackUp"


Dim objFS
Set objFS = WScript.CreateObject("Scripting.FileSystemObject")

'***BackUpフォルダを置くフォルダを作成***
If (objFS.FolderExists(motoBKFolderName) = 0) Then
	objFS.CreateFolder motoBKFolderName
End If


'***BackUp用フォルダ名を作る***'
Dim fYear,fMonth,fDay,fHour,fMinute,fSecond
fYear   = Year(Now)
fMonth  = Month(Now)
fDay    = Day(Now)
fHour   = Hour(Now)
fMinute = Minute(Now)
fSecond = Second(Now)
If (fYear < 10) Then fYear = "0" & fYear
If (fMonth < 10) Then fMonth = "0" & fMonth
If (fDay < 10) Then fDay = "0" & fDay
If (fHour < 10) Then fHour = "0" & fHour
If (fMinute < 10) Then fMinute = "0" & fMinute
If (fSecond < 10) Then fSecond = "0" & fSecond

Dim bkFolderName
bkFolderName = Right(fYear,2) & "-" & fMonth & "-" & fDay & "-" & fHour & fMinute & "-" & fSecond
bkFolderName = motoBKFolderName & "\" & bkFolderName

If (objFS.FolderExists(bkFolderName) = 0) Then
	objFS.CreateFolder bkFolderName
Else
	MsgBox "フォルダ["&bkFolderName&"]は既に存在します！！"
	WScript.Quit
End If


''***ファイルが置かれているフォルダ内のFilesオブジェクトを取得
'Dim objFolder,colFiles
'Set objFolder = objFS.GetFolder("C:\mbrnet")
'Set colFiles = objFolder.Files
'
'
''***バックアップを取る***
'Dim x
'For Each x in colFiles
'	If (x.Name <> WScript.ScriptName) Then
'		x.Copy bkFolderName&"\"&x.Name
'	End If
'Next

'***xcopyでコピーする***
xcopy "C:\mbrnet", bkFolderName




'***フォルダを圧縮する***
compressFolder motoBKFolderName




'MsgBox "バックアップフォルダ["&bkFolderName&"]を作りました！！",,Now




' コマンドプロンプトを使って圧縮する関数
Sub compressFolder(folderName)
	Dim WshShell
	Set WshShell = WScript.CreateObject("WScript.Shell")

	Dim objFS,objFolder
	Set objFS = WScript.CreateObject("Scripting.FileSystemObject")
	Set objFolder = objFS.GetFolder(folderName)

	WshShell.Run "cmd /c COMPACT /C /S:"""&objFolder.path&"""",0,1
End Sub



' コマンドプロンプトを使ってxcopyする関数
Sub xcopy(motoFolder, sakiFolder)
	Dim WshShell
	Set WshShell = WScript.CreateObject("WScript.Shell")
	
	WshShell.Run "cmd /c xcopy /I /C /E /Y /Q /R "&motoFolder&" "&sakiFolder,0,1
End Sub

