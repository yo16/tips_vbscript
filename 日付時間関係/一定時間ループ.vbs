Option Explicit

' ����̏��������I�ɉ񂷏���
' 2017/2/22 (c) y.ikeda

' �Ăяo���Ԋu�i���j
Dim intervalTime : intervalTime = 60

' ������VBS
Dim targetScript
targetScript = "��莞�ԃ��[�v_�Ă΂��.vbs"





' �҂����ԁims�j
Dim intTime_ms : intTime_ms = intervalTime * 1000 * 60

' �N���m�F
Dim retMsg
retMsg = MsgBox( intervalTime & "���Ԋu�ŁA�N�����܂��B", vbYesNo, "��莞�ԏ���")
If ( retMsg = vbNo ) Then
	WScript.Quit
End If

Dim i
Dim objShell
Set objShell = WScript.CreateObject("WScript.Shell")
For i=0 to 10	' 10��Œ�
	' �N��
	objShell.Run targetScript, 0, 1
	' �E�F�C�g
	WScript.Sleep intTime_ms
Next

