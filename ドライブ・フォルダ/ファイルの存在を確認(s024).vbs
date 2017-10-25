Option Explicit

Dim objFS

' FileSystemObject オブジェクトを生成する
Set objFS = WScript.CreateObject("Scripting.FileSystemObject")

' ファイルが存在するかどうかを表示する
' 存在する場合はTrue しない場合はFalse
WScript.Echo objFS.FileExists("abc.txt")


