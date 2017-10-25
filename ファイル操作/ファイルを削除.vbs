Option Explicit

' 2008/01/17 ikeda
' ファイルを削除
Dim objFS
Set objFS = CreateObject("Scripting.FileSystemObject")

If ( Not objFS.FileExists("削除") )Then
	msgBox "テスト用ファイル 削除.txt がありません"
	WScript.Quit
End If

'objFS.deleteFile "削除.txt"



' 読み取り専用も削除するか確認
' → 読み取り専用はNG！

' 読み取り専用対策
Dim objFile
Set objFile = objFS.GetFile( "削除.txt" )

' FileオブジェクトのAttributesプロパティを変更する
' 読み取り専用は、2ビット目
If ( objFile.Attributes and 1 ) Then
	' 読み取り専用フラグが立っていたら倒す
	objFile.Attributes = objFile.Attributes - 1
End If

objFS.deleteFile "削除.txt"
