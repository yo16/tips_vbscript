Option Explicit
' �t�@�C�����R�s�[�Q
' ���݂��Z�t�H���_���w�肵���ꍇ�A����ɍ쐬���Ă���邩�H

Dim objFS
Set objFS = WScript.CreateObject("Scripting.FileSystemObject")

'objFS.CopyFile "abc.txt",".\a\b\c\�R�s�`abc.txt"
' ���G���[�E�E�E

CreateFolder ".\a\b\c\�R�s�`abc.txt"
CreateFolder ".\a\b\c\d\e\"



' �p�X�̉��܂Ńt�H���_���쐬����
' �Ō�̕�����\�̏ꍇ�́A    �Ō�̃g�[�N�����t�H���_�Ƃ݂Ȃ�
'             \�łȂ��ꍇ�́A�Ō�̃g�[�N�����t�@�C���Ƃ݂Ȃ�
Sub CreateFolder( strInputPath )
	Dim aryDir
	aryDir = Split(strInputPath, "\")
	Dim nLastIndex
	If( Right(strInputPath,1) = "\" ) Then
		' �Ō�̓t�H���_
		nLastIndex = UBound(aryDir)
	Else
		' �Ō�̓t�@�C��
		nLastIndex = UBound(aryDir) - 1
	End If
	Dim strPath
	strPath = "."
	Dim i
	For i = 0 to nLastIndex
		strPath = strPath & "\" & aryDir(i)
		msgbox strPath
		If( objFS.FolderExists(strPath) = 0 )Then
			objFS.CreateFolder(strPath)
		End If
	Next
End Sub

