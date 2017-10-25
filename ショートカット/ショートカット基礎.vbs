Option Explicit
Dim objWshShell, objShortcut
Dim strDesktopPath
' WshShellオブジェクトを生成する
Set objWshShell = WScript.CreateObject("WScript.Shell")
' WshShortcutオブジェクトを生成する
Set objShortcut = objWshShell.CreateShortcut("test.lnk")
' ショートカットのターゲットファイルを指定する
objShortcut.TargetPath = "c:\test"
' ショートカットを作成する
objShortcut.Save
