Option Explicit
' �R�s�[
' ��������t�@�C���֏o�͂��āA
' �t�@�C������͂Ƃ���A
' �@�@�@clip < �t�@�C���p�X
' ���g���āA�R�s�[����B
' �������������t�@�C���͏����̂��x�^�[�B
' �ł����̃X�N���v�g�ł͂킩��₷���悤�Ɏc���Ă����B


Dim strTmpFilePath
strTmpFilePath = ".\copy_source.txt"

Dim objFs
Set objFs = WScript.CreateObject("Scripting.FileSystemObject")
Dim objHdl
Set objHdl = objFs.OpenTextFile( strTmpFilePath, 2, True)
objHdl.Write "copied text file!!!!"
objHdl.Close

Dim wshShell
Set wshShell = WScript.CreateObject("WScript.Shell")
Call wshShell.Run( "cmd.exe /c clip < """ & strTmpFilePath & """", 0, True )

' �ł���Ώ���
'objFs.DeleteFile strTmpFilePath

Set objFs = Nothing
Set objHdl = Nothing
Set wshShell = Nothing

msgbox "�R�s�[�������I"



