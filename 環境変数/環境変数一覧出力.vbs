Option Explicit


' ���ϐ��擾����
Dim WSHShell, WSHEnv, strEnv
Set WSHShell = WScript.CreateObject("WScript.Shell")
'      Set WSHEnv = WshShell.Environment("PROCESS")
'      Set WSHEnv = WshShell.Environment("System")
Set WSHEnv = WshShell.Environment("User")


' �t�@�C���o�͏���
Dim fileName
fileName = "EnvOut.txt"
Dim overWrite
overWrite = True

Dim objFS,objTS
Set objFS = WScript.CreateObject("Scripting.FileSystemObject")
Set objTS = objFS.CreateTextFile( fileName, overWrite )


' �t�@�C���o��
For Each strEnv In WSHEnv	' ���ׂĂ̊��ϐ����
	objTS.WriteLine strEnv
Next


' �t�@�C���N���[�Y
objTS.Close


MsgBox("���ϐ��̏o�͏I��")
