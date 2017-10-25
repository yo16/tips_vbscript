Option Explicit


Dim objFs
Set objFs = CreateObject("Scripting.FileSystemObject")
Dim objShell
Set objShell = CreateObject("WScript.Shell")

' �J�����g�t�H���_�����̃X�N���v�g������t�H���_�ɂ���
' (Drag&Drop�����Ƃ��ɕς���Ă��܂��Ή�)
objShell.CurrentDirectory = objFs.GetParentFolderName(WScript.ScriptFullName)

' ��������A�����Ώۂ̃t�H���_���擾
If WScript.Arguments.Count = 0 Then
	WScript.Exit
End If
Dim targetDir
targetDir = WScript.Arguments(0)




Dim dicExt
Set dicExt = CreateObject("Scripting.Dictionary")

' ���s
CalcExtCount targetDir

Dim msg
msg = ""
Dim key
For Each key In dicExt
	msg = msg & key & ":" & dicExt.Item(key) & vbCrLf
Next
msgbox msg




Sub CalcExtCount( dirPath )
	Dim objDir
	Set objDir = objFs.GetFolder(dirPath)
	
	Dim objFile
	Dim ext
	For Each objFile In objDir.Files
		ext = Mid(objFile.Name, InStr(objFile.Name, ".")+1)
		If dicExt.Exists(ext) Then
			dicExt.Item(ext) = dicExt.Item(ext) + 1
		Else
			dicExt.Add ext, 1
		End If
	Next
	
	Dim objSubDir
	For Each objSubDir In objDir.SubFolders
		CalcExtCount objSubDir.Path
	Next
	
End Sub

