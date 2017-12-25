' カレントフォルダを、このスクリプトが置いてあるフォルダへ変更
' 2017/03/03 (c) yo16

Dim objFS,objFolder
Set objFS = CreateObject("Scripting.FileSystemObject")
Set objFolder = objFS.GetFolder(".")
msgbox objFolder.Path	' 確認用



' スクリプトのパス
msgbox WScript.ScriptFullName	' 確認用

' フォルダパスを抽出
Dim scriptDir
scriptDir = objFS.getParentFolderName( WScript.ScriptFullName )
msgbox scriptDir	' 確認用



' カレントフォルダを変更
Dim objWshShell
Set objWshShell = WScript.CreateObject("WScript.Shell")
objWshShell.CurrentDirectory = scriptDir




Set objFolder = objFS.GetFolder(".")
msgbox objFolder.Path	' 確認用

