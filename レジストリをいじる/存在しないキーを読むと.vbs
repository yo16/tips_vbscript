Option Explicit

' 存在しないキーを読むとシステムエラー。。
' どうしたらキャッチできるかな。
' 2006/12/21 ikeda


Dim objWshShell
Set objWshShell = WScript.CreateObject("WScript.Shell")

Dim regPath
regPath = "HKLM\SOFTWARE\Microsoft\Windows\abbbcccc"

Dim tmpStr
tmpStr = objWshShell.RegRead(regPath)

msgbox "[" & tmpStr & "]"

