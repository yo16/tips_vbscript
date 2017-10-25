Option Explicit
Dim objWshShell, objShortcut
Dim strDesktopPath
' WshShellオブジェクトを生成する
Set objWshShell = WScript.CreateObject("WScript.Shell")
' WshShortcutオブジェクトを生成する
Set objShortcut = objWshShell.CreateShortcut("ip.lnk")
' ショートカットのターゲットファイルを指定する
objShortcut.TargetPath = "c:\Documents and Settings\Administrator\デスクトップ\ipmsg32_142\IPMSG.EXE"
' ショートカットを作成する
objShortcut.Save
