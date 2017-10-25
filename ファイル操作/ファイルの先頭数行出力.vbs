Option Explicit
' �t�@�C���̐擪���s���o��
' 2017 (c) y.ikeda

Dim vbsTitle : vbsTitle = "head extract"

If WScript.Arguments.Count < 1 Then
	MsgBox "�Ώۃt�@�C����Drag & Drop���Ă�������", vbOkOnly , vbsTitle
	WScript.Quit 0
End If

' �h���b�v�������ׂẴt�@�C�����A�J��Ԃ�
Dim i
For i = 0 To WScript.Arguments.Count-1
	MakeHeadFile WScript.Arguments(i)
Next


' �w�肵���t�@�C���̐擪�s���o�͂���
Sub MakeHeadFile(inFilePath)
	'MsgBox inFilePath
	
	' �o�͂���s��
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
