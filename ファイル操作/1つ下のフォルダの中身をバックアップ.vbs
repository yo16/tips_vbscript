Option Explicit

	Dim objFS
	Set objFS = WScript.CreateObject("Scripting.FileSystemObject")

	Dim objFolder,colSubFolders
	Set objFolder = objFS.GetFolder(".")
	Set colSubFolders = objFolder.SubFolders

	'***�o�b�N�A�b�v�����***
	Dim x
	For Each x in colSubFolders
		bkPlease(x.Name)
	Next


	msgbox "�I���I",,"OK!"



Sub bkPlease(currentFolderName)

	currentFolderName = currentFolderName & "\"

	'Dim YNmodori
	'YNmodori = MsgBox("�o�b�N�A�b�v������Ă������ł����H",4,"�t�H���_���ƃo�b�N�A�b�v")
	'If (YNmodori <> 6) Then
	'	WScript.Quit
	'End If




	Dim objFS
	Set objFS = WScript.CreateObject("Scripting.FileSystemObject")


	'***BackUp�p�t�H���_�������***'
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
		MsgBox "�t�H���_["&bkFolderName&"]�͊��ɑ��݂��܂��I�I"
		WScript.Quit
	End If


	Dim objFolder,colFiles
	Set objFolder = objFS.GetFolder(currentFolderName)
	Set colFiles = objFolder.Files


	'***�o�b�N�A�b�v�����***
	Dim x
	For Each x in colFiles
		If (x.Name <> WScript.ScriptName) Then
			x.Copy bkFolderName&"\"&x.Name
		End If
	Next




	'MsgBox "�o�b�N�A�b�v�t�H���_["&bkFolderName&"]�����܂����I�I",,Now


End Sub


