Option Explicit

Dim searchMsg
searchMsg = InputBox("�������镶�������͂��Ă��������B")
Dim replaceMsg
replaceMsg = InputBox("�u�����镶�������͂��Ă��������B")


Dim objFS,objFolder,colFiles
Dim x
' FileSystemObject �I�u�W�F�N�g�𐶐�����
Set objFS = WScript.CreateObject("Scripting.FileSystemObject")
' �J�����g�t�H���_��Folder�I�u�W�F�N�g���擾����
Set objFolder = objFS.GetFolder(".")
' �J�����g�t�H���_�Ɋ܂܂�邷�ׂẴt�@�C�����擾����
Set colFiles = objFolder.Files


Dim debugStr


Dim objTS
For Each x in colFiles
	debugStr = ""
	Set objTS = x.OpenAsTextStream
	Do Until objTS.AtEndOfStream
		debugStr = debugStr & objTS.ReadLine & VBCrLf
	Loop
	msgbox debugStr
Next


