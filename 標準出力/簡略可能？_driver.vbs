
Option Explicit


Dim WshShell
Set WshShell = WScript.CreateObject("WScript.Shell")

Dim intErrCode
'intErrCode=WshShell.Run("wscript �ȗ��\�H.vbs > �ȗ�����.txt",0,True)
intErrCode=WshShell.Run("cscript �ȗ��\�H.vbs > �ȗ�����.txt",0,True)


