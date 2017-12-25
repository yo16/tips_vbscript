' ショートカット作成を管理者モードで起動する
' 2016/3/15 yo16

Option Explicit

Dim objShell
Set objShell = CreateObject("Shell.Application")
objShell.ShellExecute "cscript.exe", "C:\zProgramming\VBScript\source\練習ソース\wshShellオブジェクト\ショートカット作成.vbs", "uac", "runas"
WScript.Quit
