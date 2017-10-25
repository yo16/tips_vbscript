Option Explicit

Dim WshShell
Set WshShell = WScript.CreateObject("WScript.Shell")

'第一引数:exeファイル名
'第二引数:１→可視 ０→不可視
'第三引数:１→終了を待つ ０→待たずに次を実行
'戻り値  :０→正常終了 １→異常終了
msgbox WshShell.Run("sample.bat",1,1)


