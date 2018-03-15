Option Explicit

' フォルダを移動
' 2018/3/16 yo16

Dim objFs
Set objFs = WScript.CreateObject("Scripting.FileSystemObject")

' movetestというフォルダを、aというフォルダ以下に移動する
objFs.MoveFolder ".\movetest", ".\a\"

Set objFs = Nothing
