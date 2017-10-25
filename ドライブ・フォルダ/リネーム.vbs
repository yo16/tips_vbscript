' リネーム
' 2016/5/12 y.ikeda
'
' リネーム_1.txt⇔リネーム_2.txtを切り替える
' 既にある場合は上書き
'
Option Explicit



Dim file1, file2
file1 = "リネーム_1.txt"
file2 = "リネーム_2.txt"




Dim objFs
Set objFs = CreateObject("Scripting.FileSystemObject")

Dim objFromFile

If ( objFs.FileExists( file1 ) ) Then
	fileRename objFs, file1, file2, true
Else
	If ( objFs.FileExists( file2 ) ) Then
		fileRename objFs, file2, file1, true
	End If
End If


Sub fileRename( objfs, fileFrom, fileTo, overwrite )
	If ( Not objfs.FileExists( fileFrom ) ) Then
		MsgBox "Fromファイルが存在しません"
		Exit Sub
	End If
	If ( objfs.FileExists( fileTo ) ) Then
		If ( overwrite ) Then
			objfs.DeleteFile fileTo
		Else
			MsgBox "Toファイルが既に存在します"
			Exit Sub
		End If
	End If
	
	' リネーム実行
	Dim objFile
	Set objFile = objfs.GetFile( fileFrom )
	objFile.Name = fileTo
End Sub



Set objFs = Nothing
Set objFromFile = Nothing

msgbox "end"
