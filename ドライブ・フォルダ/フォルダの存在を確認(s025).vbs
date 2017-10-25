Option Explicit

Dim objFS

' FileSystemObject オブジェクトを生成する
Set objFS = WScript.CreateObject("Scripting.FileSystemObject")

' フォルダが存在するかどうかを表示する
'存在する場合は		-1
'存在しない場合は		0
'	を返す
WScript.Echo objFS.FolderExists("c:\WINDOWS")


