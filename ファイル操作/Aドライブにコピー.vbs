Option Explicit

Dim objFS        ' FileSystemObject '
Dim objFile      ' FileObject       '
Dim objFileName  ' FileName         '
Dim copyFileName ' CopyFileName     '
Dim i            ' LoopIndex        '
Dim objDriver    ' DriverObject     '

Set objFS = CreateObject("Scripting.FileSystemObject")

' A�h���C�u�`�F�b�N!
Set objDriver = objFS.GetDrive("A")
If Not (objDriver.IsReady) Then
	MsgBox "A�h���C�u�̏������ł��Ă܂���I�I",0,"(���Q���j"
End If

For i = 0 To WScript.Arguments.Count-1
'	�t�@�C�������擾
	objFileName = WScript.Arguments(i)

'	�t�@�C�����擾
	Set objFile = objFS.GetFile(objFileName)

'	A�h���C�u�̃t�@�C���������ꍇ�̓G���[�I��
	If (objFile.Drive = "A:") Then
		MsgBox "A�h���C�u����̓_���ł��I�I",0,"(���Q���j"
	End If

'	�R�s�[��t�@�C�������쐬
	copyFileName = "A:\"&objFile.Name

'	A:\�̉��փR�s�[
	objFS.CopyFile objFileName,copyFileName
Next



