
      Dim WSHShell,WSHEnv,strList,strEnv

      Set WSHShell = WScript.CreateObject("WScript.Shell")
      Set WSHEnv = WshShell.Environment("PROCESS")
      									'1.WshEnvironmentオブジェクトを作成
MSGBOX WSHEnv.Item("USERNAME")
