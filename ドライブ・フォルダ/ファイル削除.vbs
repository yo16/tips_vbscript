' ファイルを作って消す
' 2016/5/12 y.ikeda

Option Explicit


Dim objFs
Set objFs = CreateObject("Scripting.FileSystemObject")

Dim fileName
fileName = "ファイル削除.txt"


' 準備：作成
Dim objTs
Set objTs = objFs.CreateTextFile(fileName)
objTs.WriteLine "消える運命"
objTs.Close
Set objTs = Nothing


MsgBox fileName & "を消します！"


' ファイルを削除する
objFs.DeleteFile fileName


MsgBox "end"
