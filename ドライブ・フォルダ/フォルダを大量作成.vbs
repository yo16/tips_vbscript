Option Explicit
On Error Resume Next

Dim newFolderName
newFolderName = "�V�����t�H���_��"

Dim createFolderCount
createFolderCount = InputBox("���t�H���_�̐�����͂��Ă��������B","��������́[�I")

createFolderCount = CInt(createFolderCount)
If Err Then
	MsgBox "�������Č������̂ɁB�B"
	WScript.Quit
End If
If (createFolderCount <= 0) Then
	MsgBox "������邵�Ȃ��ŁB�B"
	WScript.Quit
End If

Dim objFS
Set objFS = WScript.CreateObject("Scripting.FileSystemObject")

Dim createEndCount,loopIdx
createEndCount = 0
loopIdx = 1
While (createEndCount <> createFolderCount)
	If Not (objFS.FolderExists(newFolderName & loopIdx)) Then
		objFS.CreateFolder newFolderName & loopIdx
		createEndCount = createEndCount + 1
	End If
	loopIdx = loopIdx + 1
Wend

MsgBox "�I���`��",,"�I���`��"




















