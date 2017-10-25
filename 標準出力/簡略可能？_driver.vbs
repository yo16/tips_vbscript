
Option Explicit


Dim WshShell
Set WshShell = WScript.CreateObject("WScript.Shell")

Dim intErrCode
'intErrCode=WshShell.Run("wscript 簡略可能？.vbs > 簡略すぎ.txt",0,True)
intErrCode=WshShell.Run("cscript 簡略可能？.vbs > 簡略すぎ.txt",0,True)


