Option Explicit

Dim rtnCode
rtnCode = deleteEnvironment()



'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'�֐� deleteEnvironment
' ����		�Ȃ�
' �߂�l		����I���F0
'			�ُ�I���F-1
'
'������������
' �E���ݒ�l��ǂݍ��݁A
'   �uLOGONSERVER�v�ȊO�̊��ݒ�l���폜����
'
'2001/01/11 ikeda
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function deleteEnvironment()
	Const vbTextCompare = 1

	Dim environmentProperty
	environmentProperty = "VOLATILE"

	
	VIEW_ENV(environmentProperty)

	Dim WshShell
	Set WshShell = WScript.CreateObject("WScript.Shell")
	Dim rtnAnswer
	rtnAnswer = WshShell.Popup("�uLOGONSERVER�v�ȊO�̊��ϐ���"&vbCrLf&"�폜���Ă���낵���ł����H",0,"deleteEnvironment.vbs",36)

	if (rtnAnswer = 7) then
		Exit Function
	end if

	Dim WshEnv
	Set WshEnv = WshShell.Environment(environmentProperty)
	Dim strEnv,eqArray,deleteCount
	deleteCount = 0
	For Each strEnv In WshEnv
		eqArray = Split(strEnv,"=",-1,vbTextCompare)
		if (eqArray(0) <> "LOGONSERVER") then
			WshEnv.Remove eqArray(0)
			deleteCount = deleteCount + 1
		end if
	Next

	MsgBox deleteCount&"�̊��ϐ����폜���܂����B"

	rtnAnswer = WshShell.Popup("���ϐ������܂����H",0,"deleteEnvironment.vbs",292)
	if (rtnAnswer = 6) then
		VIEW_ENV(environmentProperty)
	end if

End Function



''' �Z�b�g���ꂽ���ϐ�������
SUB VIEW_ENV(P_EnvironmentProperty)
	Dim WSHShell2,WSHEnv2,strList,strEnv
	Set WSHShell2 = WScript.CreateObject("WScript.Shell")
	Set WSHEnv2 = WshShell2.Environment(P_EnvironmentProperty)
							'1.WshEnvironment�I�u�W�F�N�g���쐬
	MsgBox "���ϐ��̑����́A" & WSHEnv2.Count & "�ł��B"
							'2.���ϐ��̑�����\��
	strList="���ϐ��ꗗ�͈ȉ��̒ʂ�ł��B" & vbCrLf & vbCrLf
	For Each strEnv In WSHEnv2
							'3.���ׂĂ̊��ϐ����
		strList=strList & strEnv & vbCrLf
	Next
	MsgBox strList
END SUB


