Option Explicit
'
' Execでコマンドを実行して
' その結果を利用する


Dim objShell
Set objShell = WScript.CreateObject("WScript.Shell")

Dim objExec
Set objExec = objShell.Exec("cmd /c dir run*")

Dim strLine
Dim strMsg
strMsg = ""
Do Until objExec.stdout.AtEndOfStream
	strLine = objExec.stdout.ReadLine
	
	strMsg = strMsg & strLine & vbCrLf
Loop

msgbox strMsg

