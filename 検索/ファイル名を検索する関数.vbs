Option Explicit


Dim startTime
startTime = Time


Dim objFS
Dim strLine, strTemp

Set objFS = WScript.CreateObject("Scripting.FileSystemObject")

Dim objTSread,objTSwrite
'---  �T���t�@�C����������Ă���t�@�C��
Set objTSread = objFS.OpenTextFile("���݃`�F�b�N.dat",1)
'---  �������ʂ��������t�@�C��
Set objTSwrite = objFS.CreateTextFile("��������.txt",True)

strLine = ""
Dim searchFolderName
'---  ��������f�B���N�g��(�T�u�t�H���_���T��)
searchFolderName = "F:\Prsmhome\"

Do Until objTSread.AtEndOfStream
	strTemp = objTSread.ReadLine
	If Not(strTemp = "") Then
		objTSwrite.WriteLine strTemp&","&searchFile(strTemp,searchFolderName)
	End If
Loop

Dim endTime
endTime = Time
objTSwrite.WriteLine "startTime:" & startTime & "  endTime:" & endTime


objTSread.Close
objTSwrite.Close

'MsgBox "�I�������[�B"


'''''''''''''''''''''''''''''''''''''''''''''
'�֐�:searchFile
'����    P_FileName:��������t�@�C����
'        P_FolderName:������̃t�H���_��(�Ō��\�}�[�N���K�v)
'               ��΃p�X�Ńt�H���_���w��
'               �T�u�t�H���_����������(�ċA)
'               �Ȃ��ꍇ�͑SLocal�h���C�u����������<<<<<<<���쐬
'�߂�l  �t�@�C����(��΃p�X)
'               �������݂���ꍇ�ł��A
'               �͂��߂Ɍ������t�@�C���̂ݕԂ�
'''''''''''''''''''''''''''''''''''''''''''''
Function searchFile(P_FileName,P_FolderName)

	Dim objFS
	Set objFS = WScript.CreateObject("Scripting.FileSystemObject")

	'�t�@�C���̑��݃`�F�b�N
	If objFS.FileExists(P_FolderName & P_FileName) Then
		searchFile = P_FolderName & P_FileName
		Exit Function
	End If

	Dim objFolder
	Set objFolder = objFS.GetFolder(P_FolderName)

	Dim objSubFolders,objSubFolder
	Set objSubFolders = objFolder.SubFolders

	Dim rtnCode
	For Each objSubFolder In objSubFolders
		rtnCode = searchFile(P_FileName,P_FolderName&objSubFolder.Name&"\")
		If (rtnCode <> "") Then
			searchFile = rtnCode
			Exit Function
		End If
	Next

	searchFile = ""
End Function

