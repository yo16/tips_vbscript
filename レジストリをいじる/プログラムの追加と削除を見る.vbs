Option Explicit

' プログラムの追加と削除の名前一覧を出してみる


Dim objWshShell
Set objWshShell = WScript.CreateObject("WScript.Shell")

Dim regPath
regPath = "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"

Dim colEventSource
'colEventSource = objWshShell.RegRead(regPath)


'WScript.Echo colEventSource

Dim strSrc

For Each strSrc In objWshShell.RegRead(regPath)
    MsgBox "test"
    WScript.Echo strSrc
Next 



' だめだーーー
' 一覧がほしいのに。。2006/08/222006/08/222006/08/222006/08/222006/08/22
