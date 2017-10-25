Option Explicit

Dim WshShell
Set WshShell = WScript.CreateObject("WScript.Shell")

'第一引数:exeファイル名
'第二引数:１→可視 ０→不可視
'第三引数:１→終了を待つ ０→待たずに次を実行
'戻り値  :０→正常終了 １→異常終了
MSGBOX WshShell.Run("C:\winnt\system32\cmd.exe",1,1)


