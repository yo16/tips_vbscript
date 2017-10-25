' フォルダ削除
' 2015/4/24

Option Explicit

Dim objFs
Set objFs = CreateObject("Scripting.FileSystemObject")

Dim targetFolderName
targetFolderName = "delTest"

' テスト用フォルダを作成
If( Not objFs.FolderExists(targetFolderName) )Then
	objFs.CreateFolder targetFolderName
End If
If( Not objFs.FolderExists(targetFolderName & "\subFolder1") )Then
	objFs.CreateFolder targetFolderName & "\subFolder1"
End If

MsgBox targetFolderName & "を削除します！"


objFs.DeleteFolder targetFolderName
' → 中にファイルやフォルダがあっても、無条件に削除される

