Option Explicit
' ファイルオブジェクトのプロパティ

Dim objFs
Set objFs = CreateObject("Scripting.FileSystemObject")

Dim objFile
Set objFile = objFs.GetFile("a.txt")

MsgBox objFile.Name ,vbOkOnly,"ファイル名"
MsgBox objFile.Path ,vbOkOnly,"ファイルパス"
MsgBox objFile.ParentFolder ,vbOkOnly,"ParentFolder"
MsgBox objFile.DateCreated ,vbOkOnly,"作成日時"
MsgBox objFile.DateLastModified ,vbOkOnly,"更新日時"

Set objFile = Nothing
Set objFs = Nothing
