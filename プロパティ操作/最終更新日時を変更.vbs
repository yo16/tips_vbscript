Option Explicit

Dim dtmModifyDate
dtmModifyDate = CDate("2011/01/23 12:34:56")
' dtmModifyDate = CDate("2011/11/11 11:11:11")

Dim fileName
fileName = "sample.txt"
Dim folderName
folderName = "C:\zProgramming\VBScript\source\���K�\�[�X\�v���p�e�B����"

Dim objShell
Set objShell = WScript.CreateObject("Shell.Application")
Dim objFolder
Set objFolder = objShell.NameSpace(folderName)
Dim objFile
Set objFile = objFolder.ParseName(fileName)

objFile.ModifyDate = dtmModifyDate

' �Ȃ��������ɂ͔��f����Ȃ��B�B
' 11:11:11�ɐݒ肷��ƁA11:11:12�ƕ\�������B�ݒ�l��11:11:11�B
WScript.Echo "�X�V����:" & objFile.ModifyDate

Set objFile = Nothing
Set objFolder = Nothing
Set objShell = Nothing
