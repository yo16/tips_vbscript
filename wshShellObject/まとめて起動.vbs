'�܂Ƃ߂ċN��
'
'�e�L�X�g�ɏ�����Ă���c�[��(VBScript)��
'��C�ɋN�����鉡���c�[���B
'�o�b�N�A�b�v�p�ɊJ���B
'
Option Explicit


Dim matomeTXT
matomeTXT = "�܂Ƃ߂ċN���ꗗ.txt"




Dim WshShell
Set WshShell = WScript.CreateObject("WScript.Shell")


Dim objFS, objTS
Set objFS = WScript.CreateObject("Scripting.FileSystemObject")
Set objTS = objFS.OpenTextFile(matomeTXT,1)

Dim strTemp
Do Until objTS.AtEndOfStream
	strTemp = objTS.ReadLine
	If Not(strTemp = "") Then
		WshShell.Run strTemp,0,1
	End If
Loop

objTS.Close

