'CALL SET_ENV()
CALL VIEW_ENV()
'CALL DELETE_ENV()
'CALL VIEW_ENV()

''' 環境変数をセット
SUB SET_ENV()
	Set WshShell = WScript.CreateObject("WScript.Shell")
	Set WshEnv = WshShell.Environment("VOLATILE")
	WshEnv("test")="hoge"
				'1."test"という環境変数に、"hoge"という値をセット
	MsgBox "環境変数testを定義しました。"
END SUB



''' 環境変数 "test" を削除
SUB DELETE_ENV()
	Set WshShell3 = WScript.CreateObject("WScript.Shell")
	Set WshEnv3 = WshShell3.Environment("VOLATILE")
	WshEnv3.Remove "test"
				'2.test"という環境変数を削除
	MsgBox "環境変数testを削除しました。"
END SUB


''' セットされた環境変数を見る
SUB VIEW_ENV()
	Dim WSHShell2,WSHEnv2,strList,strEnv
	Set WSHShell2 = WScript.CreateObject("WScript.Shell")
	Set WSHEnv2 = WshShell2.Environment("VOLATILE")
							'1.WshEnvironmentオブジェクトを作成
	MsgBox "環境変数の総数は、" & WSHEnv2.Count & "です。"
							'2.環境変数の総数を表示
	strList="環境変数一覧は以下の通りです。" & vbCrLf & vbCrLf
	For Each strEnv In WSHEnv2
							'3.すべての環境変数を列挙
		strList=strList & strEnv & vbCrLf
	Next
	MsgBox strList
END SUB

