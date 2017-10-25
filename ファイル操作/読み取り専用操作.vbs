'-----------------------------------
' 読み取り専用プロパティを変更する
'-----------------------------------
Option Explicit

Dim FileName
FileName = "readonly.txt"

Dim objFS
Set objFS = CreateObject("Scripting.FileSystemObject")
Dim objFile
Set objFile = objFS.GetFile( FileName )

' FileオブジェクトのAttributesプロパティを変更する
' 読み取り専用は、2ビット目
If ( objFile.Attributes and 1 ) Then
	' 読み取り専用フラグが立っていたら倒す
	objFile.Attributes = objFile.Attributes - 1
Else
	' 読み取り専用フラグが倒れていたら立てる
	objFile.Attributes = objFile.Attributes + 1
End If

