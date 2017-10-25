Option Explicit

	Dim objFS
	Set objFS = WScript.CreateObject("Scripting.FileSystemObject")

	Dim objFolder,colSubFolders
	Set objFolder = objFS.GetFolder(".")
	Set colSubFolders = objFolder.SubFolders

	'***バックアップを取る***
	Dim x
	For Each x in colSubFolders
		bkPlease(x.Name)
	Next


	msgbox "終了！",,"OK!"



Sub bkPlease(currentFolderName)

	currentFolderName = currentFolderName & "\"

	'Dim YNmodori
	'YNmodori = MsgBox("バックアップを取ってもいいですか？",4,"フォルダごとバックアップ")
	'If (YNmodori <> 6) Then
	'	WScript.Quit
	'End If




	Dim objFS
	Set objFS = WScript.CreateObject("Scripting.FileSystemObject")


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

	bkFolderName = currentFolderName & bkFolderName

	If (objFS.FolderExists(bkFolderName) = 0) Then
		objFS.CreateFolder bkFolderName
	Else
		MsgBox "フォルダ["&bkFolderName&"]は既に存在します！！"
		WScript.Quit
	End If


	Dim objFolder,colFiles
	Set objFolder = objFS.GetFolder(currentFolderName)
	Set colFiles = objFolder.Files


	'***バックアップを取る***
	Dim x
	For Each x in colFiles
		If (x.Name <> WScript.ScriptName) Then
			x.Copy bkFolderName&"\"&x.Name
		End If
	Next




	'MsgBox "バックアップフォルダ["&bkFolderName&"]を作りました！！",,Now


End Sub


