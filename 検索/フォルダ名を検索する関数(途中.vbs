

'''''''''''''''''''''''''''''''''''''''''''''
'�֐�:searchFolder
'����    P_FolderName:��������t�@�C����
'        P_SearchFolderPath:������̃t�H���_��
'               ��΃p�X�Ńt�H���_���w��(�Ō��\�}�[�N���K�v)
'               �T�u�t�H���_����������(�ċA)
'               �Ȃ��ꍇ�͑SLocal�h���C�u����������<<<<<���쐬
'�߂�l  �t�@�C����(��΃p�X)
'               �������݂���ꍇ�ł��A
'               �͂��߂Ɍ������t�@�C���̂ݕԂ�
'''''''''''''''''''''''''''''''''''''''''''''
Function searchFolder(P_FolderName,P_SearchFolderPath)

	Dim objFS
	Set objFS = WScript.CreateObject("Scripting.FileSystemObject")

	'�t�H���_�̑��݃`�F�b�N
	If objFS.FolderExists(P_SearchFolderPath & P_FolderName) Then
		searchFile = P_SearchFolderPath & P_FolderName
		Exit Function
	End If

	Dim objFolder
	Set objFolder = objFS.GetFolder(P_SearchFolderPath)

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

