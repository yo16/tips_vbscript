Option Explicit

' �t�@�C�����R�s�[
' 2017/5/1 y.ikeda

Dim objFs
Set objFs = WScript.CreateObject("Scripting.FileSystemObject")


Dim overwrite
overwrite = True

' from, to, overwrite(�f�t�H���g:true)
objFs.CopyFile "tree.txt", "tree2.txt", overwrite


