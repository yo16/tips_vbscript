Option Explicit
On Error Resume Next

Dim newFileName
newFileName = "新しいテキストファイル☆"
Dim kakutyoushi
kakutyoushi = ".txt"

Dim createFileCount
createFileCount = InputBox("作るファイルの数を入力してください。","整数を入力ー！")

createFileCount = CInt(createFileCount)
If Err Then
	MsgBox "整数って言ったのに。。"
	WScript.Quit
End If
If (createFileCount <= 0) Then
	MsgBox "いじわるしないで。。"
	WScript.Quit
End If

Dim objFS
Set objFS = WScript.CreateObject("Scripting.FileSystemObject")

Dim createEndCount,loopIdx
createEndCount = 0
loopIdx = 1
While (createEndCount <> createFileCount)
	If Not (objFS.FileExists(newFileName & loopIdx & kakutyoushi)) Then
		objFS.CreateTextFile(newFileName & loopIdx & kakutyoushi)
		createEndCount = createEndCount + 1
	End If
	loopIdx = loopIdx + 1
Wend

MsgBox "終了～☆",,"終了～☆"


If Err Then
	MsgBox "エラー！"
	WScript.Quit
End If


















