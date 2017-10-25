Option Explicit

Dim dtmModifyDate
dtmModifyDate = CDate("2011/01/23 12:34:56")
' dtmModifyDate = CDate("2011/11/11 11:11:11")

Dim fileName
fileName = "sample.txt"
Dim folderName
folderName = "C:\zProgramming\VBScript\source\練習ソース\プロパティ操作"

Dim objShell
Set objShell = WScript.CreateObject("Shell.Application")
Dim objFolder
Set objFolder = objShell.NameSpace(folderName)
Dim objFile
Set objFile = objFolder.ParseName(fileName)

objFile.ModifyDate = dtmModifyDate

' なぜかすぐには反映されない。。
' 11:11:11に設定すると、11:11:12と表示される。設定値は11:11:11。
WScript.Echo "更新日時:" & objFile.ModifyDate

Set objFile = Nothing
Set objFolder = Nothing
Set objShell = Nothing
