' ���l�[��
' 2016/5/12 y.ikeda
'
' ���l�[��_1.txt�̃��l�[��_2.txt��؂�ւ���
' ���ɂ���ꍇ�͏㏑��
'
Option Explicit



Dim file1, file2
file1 = "���l�[��_1.txt"
file2 = "���l�[��_2.txt"




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
		MsgBox "From�t�@�C�������݂��܂���"
		Exit Sub
	End If
	If ( objfs.FileExists( fileTo ) ) Then
		If ( overwrite ) Then
			objfs.DeleteFile fileTo
		Else
			MsgBox "To�t�@�C�������ɑ��݂��܂�"
			Exit Sub
		End If
	End If
	
	' ���l�[�����s
	Dim objFile
	Set objFile = objfs.GetFile( fileFrom )
	objFile.Name = fileTo
End Sub



Set objFs = Nothing
Set objFromFile = Nothing

msgbox "end"
