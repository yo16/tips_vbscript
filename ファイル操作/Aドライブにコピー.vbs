Option Explicit

Dim objFS        ' FileSystemObject '
Dim objFile      ' FileObject       '
Dim objFileName  ' FileName         '
Dim copyFileName ' CopyFileName     '
Dim i            ' LoopIndex        '
Dim objDriver    ' DriverObject     '

Set objFS = CreateObject("Scripting.FileSystemObject")

' Aドライブチェック!
Set objDriver = objFS.GetDrive("A")
If Not (objDriver.IsReady) Then
	MsgBox "Aドライブの準備ができてません！！",0,"(ｘ＿ｘ）"
End If

For i = 0 To WScript.Arguments.Count-1
'	ファイル名を取得
	objFileName = WScript.Arguments(i)

'	ファイルを取得
	Set objFile = objFS.GetFile(objFileName)

'	Aドライブのファイルだった場合はエラー終了
	If (objFile.Drive = "A:") Then
		MsgBox "Aドライブからはダメです！！",0,"(ｘ＿ｘ）"
	End If

'	コピー先ファイル名を作成
	copyFileName = "A:\"&objFile.Name

'	A:\の下へコピー
	objFS.CopyFile objFileName,copyFileName
Next



