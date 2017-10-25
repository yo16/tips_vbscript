
      Dim WSHShell,WSHEnv,strList,strEnv

      Set WSHShell = WScript.CreateObject("WScript.Shell")
'      Set WSHEnv = WshShell.Environment("PROCESS")
'      Set WSHEnv = WshShell.Environment("System")
      Set WSHEnv = WshShell.Environment("User")
      									'1.WshEnvironmentオブジェクトを作成

      MsgBox "Windowsインストールフォルダは、" & WSHEnv.Item("windir") & "です。" 
      									'2.Windowsがインストールされているフォルダ名を表示

      MsgBox "環境変数の総数は、" & WSHEnv.Count & "です。"
      									'3.環境変数の総数を表示

      strList="環境変数一覧は以下の通りです。" & vbCrLf
      Dim i
      i = 0
      For Each strEnv In WSHEnv
      									'4.すべての環境変数を列挙
		i = i + 1
        strList=strList & strEnv & vbCrLf
        If i > 10 Then
        	MsgBox strList
        	strList = ""
        	i = 0
        End If
      Next
      MsgBox strList