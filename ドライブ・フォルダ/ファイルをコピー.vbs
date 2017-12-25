Option Explicit

' ファイルをコピー
' 2017/5/1 yo16

Dim objFs
Set objFs = WScript.CreateObject("Scripting.FileSystemObject")


Dim overwrite
overwrite = True

' from, to, overwrite(デフォルト:true)
objFs.CopyFile "tree.txt", "tree2.txt", overwrite


