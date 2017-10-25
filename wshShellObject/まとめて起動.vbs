'まとめて起動
'
'テキストに書かれているツール(VBScript)を
'一気に起動する横着ツール。
'バックアップ用に開発。
'
Option Explicit


Dim matomeTXT
matomeTXT = "まとめて起動一覧.txt"




Dim WshShell
Set WshShell = WScript.CreateObject("WScript.Shell")


Dim objFS, objTS
Set objFS = WScript.CreateObject("Scripting.FileSystemObject")
Set objTS = objFS.OpenTextFile(matomeTXT,1)

Dim strTemp
Do Until objTS.AtEndOfStream
	strTemp = objTS.ReadLine
	If Not(strTemp = "") Then
		WshShell.Run strTemp,0,1
	End If
Loop

objTS.Close

