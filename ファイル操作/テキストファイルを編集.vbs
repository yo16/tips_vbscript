Option Explicit

'--�ҏW�t�@�C����
Dim editFileName,workFileName
editFileName = "sample2.txt"
workFileName = editFileName & ".work"

Const ForReading = 1, ForWriting = 2, ForAppending = 8
Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0

Dim objFS
Set objFS = CreateObject("Scripting.FileSystemObject")

'--�ҏW���̃t�@�C�����J��
Dim objEditFile
Set objEditFile = objFS.GetFile(editFileName)
Dim objEditTS
Set objEditTS = objEditFile.OpenAsTextStream(ForReading,TristateUseDefault)

'--�ҏW��(Work�t�@�C��)�̃t�@�C�����쐬
Dim objWorkTS
Set objWorkTS = objFS.CreateTextFile(workFileName,False)

'--�ҏW
Do Until objEditTS.AtEndOfStream
	'--���̏ꍇ��[']���n�߂ɂ��Ă���B
	objWorkTS.WriteLine "'" & objEditTS.ReadLine
Loop

'--�t�@�C�����N���[�Y
objEditTS.Close
objWorkTS.Close

'--�ҏW���̃t�@�C�����폜
objEditFile.Delete

'--�ҏW��̃t�@�C������ҏW���̃t�@�C�����ɕύX
Dim objWorkFile
Set objWorkFile = objFS.GetFile(workFileName)
objWorkFile.Name = editFileName


MsgBox "�I���`�`�`�I",,"�R(�P���P)�m"

