Option Explicit
Dim objWshShell, objShortcut
Dim strDesktopPath
Set objWshShell = WScript.CreateObject("WScript.Shell")

Dim oracleHomePath
oracleHomePath = GetOracleHome()

strDesktopPath = objWshShell.SpecialFolders("Desktop")

Set objShortcut = objWshShell.CreateShortcut(strDesktopPath & "\ソース管理.lnk")

' ショートカットのターゲットファイルを指定する
objShortcut.TargetPath = oracleHomePath & "\Bin\ifrun60.EXE"
' ショートカットに渡す引数を指定する
objShortcut.Arguments = "c:\prism\form\ファイル管理.fmx nova01/nova01@smtap.world"
' ショートカットを作成する
objShortcut.Save





Function GetOracleHome()
'---------------------------------------------------
' レジストリを参照して、ORACLE_HOMEのパスを返す関数
'---------------------------------------------------
	Dim objWshShell
	Dim RegData

	Set objWshShell = WScript.CreateObject ("WScript.Shell")
	RegData = "HKLM\Software\ORACLE\ORACLE_HOME"
	GetOracleHome = objWshShell.RegRead(RegData)

End Function
