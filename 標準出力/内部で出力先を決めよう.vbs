
Option Explicit

Wscript.Stdout = "D:\WSH\練習ソース\標準出力\変更できたぞ！.txt"


'↓この辺りを変える。
Dim stdout
'set stdout = "D:\WSH\練習ソース\標準出力\変更できたぞ！.txt"
set stdout = Wscript.Stdout




'ファイルに出力される予定のメッセージ。
WScript.Echo "変更できたかな。"





