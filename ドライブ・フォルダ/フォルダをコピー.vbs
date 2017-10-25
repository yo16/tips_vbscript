Option Explicit

' フォルダをコピー
' 2016/1/22 y.ikeda

Dim objFs
Set objFs = WScript.CreateObject("Scripting.FileSystemObject")


' from
' to
' 上書き可否
objFs.CopyFolder "a", "a_copy", True

' 上書きFalseで既にある場合は、下記msgが出る
' ---------------------------
' Windows Script Host
' ---------------------------
' スクリプト:	C:\zProgramming\VBScript\source\練習ソース\ドライブ・フォルダ\フォルダをコピー.vbs
' 行:	13
' 文字:	1
' エラー:	既に同名のファイルが存在しています。
' コード:	800A003A
' ソース: 	Microsoft VBScript 実行時エラー
' 
' ---------------------------
' OK   
' ---------------------------



msgbox "end"


