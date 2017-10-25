Option Explicit
' ファイルの先頭数行を出力
' 2017 (c) y.ikeda

Dim vbsTitle : vbsTitle = "head extract"

If WScript.Arguments.Count < 1 Then
	MsgBox "対象ファイルをDrag & Dropしてください", vbOkOnly , vbsTitle
	WScript.Quit 0
End If

' ドロップしたすべてのファイル分、繰り返す
Dim i
For i = 0 To WScript.Arguments.Count-1
	MakeHeadFile WScript.Arguments(i)
Next


' 指定したファイルの先頭行を出力する
Sub MakeHeadFile(inFilePath)
	'MsgBox inFilePath
	
	' 出力する行数
	Dim outLineNum : outLineNum = 10
	
	
	Dim outFilePath : outFilePath = inFilePath & "_out.txt"
	
	Dim objFs : Set objFs = WScript.CreateObject("Scripting.FileSystemObject")
	Dim objTs : Set objTs = objFs.OpenTextFile( inFilePath )
	Dim objOutTs : Set objOutTs = objFs.CreateTextFile( outFilePath, True ) ' true:Overwrite
	
	Do While(( outLineNum > 0 ) And (objTs.AtEndOfStream = False))
		objOutTs.WriteLine objTs.ReadLine
		
		outLineNum = outLineNum - 1
	Loop
	
	
	objOutTs.Close
	Set objOutTs = Nothing
	objTs.Close
	Set objTs = Nothing
	Set objFs = Nothing
	
	
	
End Sub
