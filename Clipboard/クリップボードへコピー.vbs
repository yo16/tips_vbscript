Option Explicit
' コピー
' いったんファイルへ出力して、
' ファイルを入力とする、
' 　　　clip < ファイルパス
' を使って、コピーする。
' いったん作ったファイルは消すのがベター。
' でもこのスクリプトではわかりやすいように残しておく。


Dim strTmpFilePath
strTmpFilePath = ".\copy_source.txt"

Dim objFs
Set objFs = WScript.CreateObject("Scripting.FileSystemObject")
Dim objHdl
Set objHdl = objFs.OpenTextFile( strTmpFilePath, 2, True)
objHdl.Write "copied text file!!!!"
objHdl.Close

Dim wshShell
Set wshShell = WScript.CreateObject("WScript.Shell")
Call wshShell.Run( "cmd.exe /c clip < """ & strTmpFilePath & """", 0, True )

' できれば消す
'objFs.DeleteFile strTmpFilePath

Set objFs = Nothing
Set objHdl = Nothing
Set wshShell = Nothing

msgbox "コピーしたぞ！"



