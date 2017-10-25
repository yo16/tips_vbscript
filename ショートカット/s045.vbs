Option Explicit
Dim objWshShell, objShortcut
Dim strDesktopPath
' WshShellオブジェクトを生成する
Set objWshShell = WScript.CreateObject("WScript.Shell")
' デスクトップのフォルダ名を取得する
strDesktopPath = objWshShell.SpecialFolders("Desktop")
' WshShortcutオブジェクトを生成する
Set objShortcut = objWshShell.CreateShortcut(strDesktopPath & "\複数の表.lnk")
' ショートカットのターゲットファイルを指定する
objShortcut.TargetPath = "c:\program files\microsoft office\office\excel.exe"
' ショートカットに渡す引数を指定する
objShortcut.Arguments = "c:\home\wsh\ch02\s1.xls c:\home\wsh\ch02\s2.xls"
' ショートカットを作成する
objShortcut.Save
