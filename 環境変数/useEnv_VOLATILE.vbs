'CALL SET_ENV()
CALL VIEW_ENV()
'CALL DELETE_ENV()
'CALL VIEW_ENV()

''' ���ϐ����Z�b�g
SUB SET_ENV()
	Set WshShell = WScript.CreateObject("WScript.Shell")
	Set WshEnv = WshShell.Environment("VOLATILE")
	WshEnv("test")="hoge"
				'1."test"�Ƃ������ϐ��ɁA"hoge"�Ƃ����l���Z�b�g
	MsgBox "���ϐ�test���`���܂����B"
END SUB



''' ���ϐ� "test" ���폜
SUB DELETE_ENV()
	Set WshShell3 = WScript.CreateObject("WScript.Shell")
	Set WshEnv3 = WshShell3.Environment("VOLATILE")
	WshEnv3.Remove "test"
				'2.test"�Ƃ������ϐ����폜
	MsgBox "���ϐ�test���폜���܂����B"
END SUB


''' �Z�b�g���ꂽ���ϐ�������
SUB VIEW_ENV()
	Dim WSHShell2,WSHEnv2,strList,strEnv
	Set WSHShell2 = WScript.CreateObject("WScript.Shell")
	Set WSHEnv2 = WshShell2.Environment("VOLATILE")
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

