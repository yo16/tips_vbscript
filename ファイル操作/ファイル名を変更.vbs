Option Explicit
' �t�@�C������ύX����
' 2008/01/16 y.ikeda


' MoveFile���g��



Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")
fso.MoveFile "rename_1.txt", "rename_2.txt"


