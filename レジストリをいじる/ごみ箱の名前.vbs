Option Explicit


'**	�J�X�^�}�C�Y���悤�I�V���[�Y
'**		���ݔ��̖��O��ς����
'**				2001/01/29



Dim objWshShell
Set objWshShell = WScript.CreateObject("WScript.Shell")

Dim regPath
regPath = "HKCR\CLSID\{645FF040-5081-101B-9F08-00AA002F954E}\"

Dim oldName
oldName = objWshShell.RegRead(regPath)

Dim newName
newName = InputBox("���ݔ���" & vbCrLf & "�V�������O��" & vbCrLf & "���Ă����܂��傤��","�i~��~���j����",oldName)

If ( (newName = "") or (newName = oldName) ) Then
	MsgBox "�ς��܂���ł����Ƃ��B�B",0,"(�~_�~; )"
Else
	MsgBox "�u" & newName & "�v�ɕς��Ƃ��܂����B" & vbCrLf & "�ă��O�I��������L���ł��B",0,"(�P�[�P)v"
	MsgBox "�����Y��܂������A",0,"������ƃI�V���Z�B(�E_�E?)"
	MsgBox "���W�X�g�����������Ă܂��B" & vbCrLf & "�Ȃ񂩂����Ă��ӔC�͕����܂���B" & vbCrLf & "���߂�ˁ`�B���傤���Ȃ���ˁ`�B",16,"�u(������) �A�C�^�b�I"
	Dim modori
	modori = MsgBox("���́u" & oldName & "�v�ɖ߂��܂����H",292,"���͂܂��Ԃɍ����B�R�i�ށ܁R�j�i�m�܁ށj�m")

	If (modori = 7) Then
		objWshShell.RegWrite regPath,newName
		MsgBox "�}�W�ŕς��܂����B" & vbCrLf & "�߂������Ƃ��͂�����x����Ăˁ`��",0,"(o��o)/~�}�^�l�F�`"
	Else
		MsgBox "�������Ȃ��B",,"o(�E_�E)9 �A�b�p�[�I"
	End If

End If



