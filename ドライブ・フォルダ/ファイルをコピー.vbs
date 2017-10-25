Option Explicit

' ファイルをコピー
' 2017/5/1 y.ikeda

Dim objFs
Set objFs = WScript.CreateObject("Scripting.FileSystemObject")


Dim overwrite
overwrite = True

' from, to, overwrite(デフォルト:true)
objFs.CopyFile "tree.txt", "tree2.txt", overwrite


