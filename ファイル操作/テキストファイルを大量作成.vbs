Option Explicit
On Error Resume Next

Dim newFileName
newFileName = "�V�����e�L�X�g�t�@�C����"
Dim kakutyoushi
kakutyoushi = ".txt"

Dim createFileCount
createFileCount = InputBox("���t�@�C���̐�����͂��Ă��������B","��������́[�I")

createFileCount = CInt(createFileCount)
If Err Then
	MsgBox "�������Č������̂ɁB�B"
	WScript.Quit
End If
If (createFileCount <= 0) Then
	MsgBox "������邵�Ȃ��ŁB�B"
	WScript.Quit
End If

Dim objFS
Set objFS = WScript.CreateObject("Scripting.FileSystemObject")

Dim createEndCount,loopIdx
createEndCount = 0
loopIdx = 1
While (createEndCount <> createFileCount)
	If Not (objFS.FileExists(newFileName & loopIdx & kakutyoushi)) Then
		objFS.CreateTextFile(newFileName & loopIdx & kakutyoushi)
		createEndCount = createEndCount + 1
	End If
	loopIdx = loopIdx + 1
Wend

MsgBox "�I���`��",,"�I���`��"


If Err Then
	MsgBox "�G���[�I"
	WScript.Quit
End If


















