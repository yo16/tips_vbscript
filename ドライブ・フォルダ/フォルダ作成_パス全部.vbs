' �w�肵���t�H���_�p�X�́A�r�����Ȃ��Ă��S�����
Option Explicit



Dim objFS
Set objFS = WScript.CreateObject("Scripting.FileSystemObject")

Dim strPath
strPath = ".\aa\bb\cccc"

Call subCreDir(strPath)


Sub subCreDir(path)
	If objFS.FolderExists(path) Then
		exit sub
	End If
	
	' ��납��\�܂Ŕ����o��(\�܂܂�)
	Dim nSepPos
	nSepPos = InStrRev(path, "\")
	Dim strDirName
	strDirName = Mid(path, nSepPos+1)
	'msgbox strDirName
	
	' �t�H���_��ʃt�H���_�̑��݃`�F�b�N
	If Not objFS.FolderExists(Left(path,nSepPos-1)) Then
		' ���݂��Ȃ��ꍇ�A��ʃt�H���_���쐬����
		Call subCreDir( Left(path,nSepPos-1) )
	End If
	
	' �J�����g�̃t�H���_���쐬����
	objFS.CreateFolder(path)
End Sub

