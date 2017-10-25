Option Explicit
On Error Resume Next

Dim newFolderName
newFolderName = "新しいフォルダ☆"

Dim createFolderCount
createFolderCount = InputBox("作るフォルダの数を入力してください。","整数を入力ー！")

createFolderCount = CInt(createFolderCount)
If Err Then
	MsgBox "整数って言ったのに。。"
	WScript.Quit
End If
If (createFolderCount <= 0) Then
	MsgBox "いじわるしないで。。"
	WScript.Quit
End If

Dim objFS
Set objFS = WScript.CreateObject("Scripting.FileSystemObject")

Dim createEndCount,loopIdx
createEndCount = 0
loopIdx = 1
While (createEndCount <> createFolderCount)
	If Not (objFS.FolderExists(newFolderName & loopIdx)) Then
		objFS.CreateFolder newFolderName & loopIdx
		createEndCount = createEndCount + 1
	End If
	loopIdx = loopIdx + 1
Wend

MsgBox "終了～☆",,"終了～☆"




















