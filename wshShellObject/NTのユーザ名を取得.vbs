
      Dim WSHShell,WSHEnv,strList,strEnv

      Set WSHShell = WScript.CreateObject("WScript.Shell")
      Set WSHEnv = WshShell.Environment("PROCESS")
      									'1.WshEnvironment�I�u�W�F�N�g���쐬
MSGBOX WSHEnv.Item("USERNAME")
