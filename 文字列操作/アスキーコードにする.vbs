Option Explicit

'�ϊ����镶�������
Dim exStr
exStr = InputBox("�A�X�L�[�R�[�h�ɕς��镶�����"&VBCrLf&"���͂��Ă݂Ă��������I","���͂��Ă݁H")


'�L�����Z��or���͂���Ă��Ȃ��ꍇ
If (exStr = "") Then WScript.Quit


'������
'MsgBox "length = " & Len(exStr)

'�P�������ϊ�
Dim idx
Dim rtnStr
rtnStr = ""
For idx = 1 to Len(exStr)
	'�P�������o��
	'MsgBox "����" & idx & " = " & Mid(exStr,idx,1)

	'�A�X�L�[�R�[�h�ɂ��Ċi�[
	rtnStr = rtnStr & "Chr(" & Asc(Mid(exStr,idx,1)) & ")&"
Next

'�Ō��[&]�����
rtnStr = Left(rtnStr,Len(rtnStr)-1)

'���s���ʏo��(MsgBox)
'MsgBox rtnStr

'���s���ʏo��(InputBox)
Dim modori
modori = InputBox("[" & exStr & "]��" & VBCrLf & "ASCII�R�[�h�ɕς��܂����I","���ʔ��\�`��",rtnStr)














