Option Explicit


'Dim YNmodori
'YNmodori = MsgBox("�o�b�N�A�b�v������Ă������ł����H",4,"�t�H���_���ƃo�b�N�A�b�v")
'If (YNmodori <> 6) Then
'	WScript.Quit
'End If


Dim motoBKFolderName
motoBKFolderName = "MartBrowserSourceBackUp"


Dim objFS
Set objFS = WScript.CreateObject("Scripting.FileSystemObject")

'***BackUp�t�H���_��u���t�H���_���쐬***
If (objFS.FolderExists(motoBKFolderName) = 0) Then
	objFS.CreateFolder motoBKFolderName
End If


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
bkFolderName = motoBKFolderName & "\" & bkFolderName

If (objFS.FolderExists(bkFolderName) = 0) Then
	objFS.CreateFolder bkFolderName
Else
	MsgBox "�t�H���_["&bkFolderName&"]�͊��ɑ��݂��܂��I�I"
	WScript.Quit
End If


''***�t�@�C�����u����Ă���t�H���_����Files�I�u�W�F�N�g���擾
'Dim objFolder,colFiles
'Set objFolder = objFS.GetFolder("C:\mbrnet")
'Set colFiles = objFolder.Files
'
'
''***�o�b�N�A�b�v�����***
'Dim x
'For Each x in colFiles
'	If (x.Name <> WScript.ScriptName) Then
'		x.Copy bkFolderName&"\"&x.Name
'	End If
'Next

'***xcopy�ŃR�s�[����***
xcopy "C:\mbrnet", bkFolderName




'***�t�H���_�����k����***
compressFolder motoBKFolderName




'MsgBox "�o�b�N�A�b�v�t�H���_["&bkFolderName&"]�����܂����I�I",,Now




' �R�}���h�v�����v�g���g���Ĉ��k����֐�
Sub compressFolder(folderName)
	Dim WshShell
	Set WshShell = WScript.CreateObject("WScript.Shell")

	Dim objFS,objFolder
	Set objFS = WScript.CreateObject("Scripting.FileSystemObject")
	Set objFolder = objFS.GetFolder(folderName)

	WshShell.Run "cmd /c COMPACT /C /S:"""&objFolder.path&"""",0,1
End Sub



' �R�}���h�v�����v�g���g����xcopy����֐�
Sub xcopy(motoFolder, sakiFolder)
	Dim WshShell
	Set WshShell = WScript.CreateObject("WScript.Shell")
	
	WshShell.Run "cmd /c xcopy /I /C /E /Y /Q /R "&motoFolder&" "&sakiFolder,0,1
End Sub

