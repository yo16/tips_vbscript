
Option Explicit

Dim str
str = "c:\test\aa\bb.txt"

Dim objFs
Set objFs = CreateObject("Scripting.FileSystemObject")

' ファイルの存在有無は無関係

msgbox objFs.GetFileName(str)
' bb.txt

msgbox objFs.GetBaseName(str)
' bb

msgbox objFs.GetParentFolderName(str)
' c:\test\aa


