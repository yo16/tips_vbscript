Option Explicit

'================================='
Dim stringSize
stringSize = 20
Dim sizeMark
sizeMark = "●"
Dim maxMark
maxMark = 30
'================================='

If (WScript.Arguments.Count = 0) Then
	MsgBox "フォルダを落として下さい。"
	WScript.Quit
End If

Dim folderName
folderName = WScript.Arguments(0)

Dim objFS,objFolder,objSubFolders
Set objFS = WScript.CreateObject("Scripting.FileSystemObject")
If Not (objFS.FolderExists(folderName)) Then
	MsgBox "ファイルを落としてもダメです。"
	WScript.Quit
End If
Set objFolder = objFS.GetFolder(folderName)
Set objSubFolders = objFolder.SubFolders

Dim fullSize,idx
fullSize = objFolder.Size

Dim rtnString
rtnString = ""

rtnString = "[[ " & Left(objFolder.Name,stringSize) & " ]]" & " "
rtnString = rtnString & "(" & fullSize & " byte)" & vbCrLf
For idx = 1 to maxMark
	rtnString = rtnString & sizeMark
Next
rtnString = rtnString & vbCrLf & vbCrLf


Dim xFolder
Dim subFolderSize,markCount
For Each xFolder In objSubFolders
	rtnString = rtnString & "[ " & Left(xFolder.Name,stringSize) & " ]" & " "
	subFolderSize = xFolder.Size
	rtnString = rtnString & "(" & subFolderSize & " byte)" & vbCrLf
	markCount = (subFolderSize * maxMark) \ fullSize
	For idx = 1 to markCount
		rtnString = rtnString & sizeMark
	Next
	rtnString = rtnString & vbCrLf
Next


MsgBox rtnString
