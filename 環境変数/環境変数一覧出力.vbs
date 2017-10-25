Option Explicit


' 環境変数取得準備
Dim WSHShell, WSHEnv, strEnv
Set WSHShell = WScript.CreateObject("WScript.Shell")
'      Set WSHEnv = WshShell.Environment("PROCESS")
'      Set WSHEnv = WshShell.Environment("System")
Set WSHEnv = WshShell.Environment("User")


' ファイル出力準備
Dim fileName
fileName = "EnvOut.txt"
Dim overWrite
overWrite = True

Dim objFS,objTS
Set objFS = WScript.CreateObject("Scripting.FileSystemObject")
Set objTS = objFS.CreateTextFile( fileName, overWrite )


' ファイル出力
For Each strEnv In WSHEnv	' すべての環境変数を列挙
	objTS.WriteLine strEnv
Next


' ファイルクローズ
objTS.Close


MsgBox("環境変数の出力終了")
