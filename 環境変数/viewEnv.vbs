
      Dim WSHShell,WSHEnv,strList,strEnv

      Set WSHShell = WScript.CreateObject("WScript.Shell")
'      Set WSHEnv = WshShell.Environment("PROCESS")
'      Set WSHEnv = WshShell.Environment("System")
      Set WSHEnv = WshShell.Environment("User")
      									'1.WshEnvironment�I�u�W�F�N�g���쐬

      MsgBox "Windows�C���X�g�[���t�H���_�́A" & WSHEnv.Item("windir") & "�ł��B" 
      									'2.Windows���C���X�g�[������Ă���t�H���_����\��

      MsgBox "���ϐ��̑����́A" & WSHEnv.Count & "�ł��B"
      									'3.���ϐ��̑�����\��

      strList="���ϐ��ꗗ�͈ȉ��̒ʂ�ł��B" & vbCrLf
      Dim i
      i = 0
      For Each strEnv In WSHEnv
      									'4.���ׂĂ̊��ϐ����
		i = i + 1
        strList=strList & strEnv & vbCrLf
        If i > 10 Then
        	MsgBox strList
        	strList = ""
        	i = 0
        End If
      Next
      MsgBox strList