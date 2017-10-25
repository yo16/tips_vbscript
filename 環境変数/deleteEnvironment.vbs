Option Explicit

Dim rtnCode
rtnCode = deleteEnvironment()



'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'関数 deleteEnvironment
' 引数		なし
' 戻り値		正常終了：0
'			異常終了：-1
'
'＊＊説明＊＊
' ・環境設定値を読み込み、
'   「LOGONSERVER」以外の環境設定値を削除する
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
	rtnAnswer = WshShell.Popup("「LOGONSERVER」以外の環境変数を"&vbCrLf&"削除してもよろしいですか？",0,"deleteEnvironment.vbs",36)

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

	MsgBox deleteCount&"個の環境変数を削除しました。"

	rtnAnswer = WshShell.Popup("環境変数を見ますか？",0,"deleteEnvironment.vbs",292)
	if (rtnAnswer = 6) then
		VIEW_ENV(environmentProperty)
	end if

End Function



''' セットされた環境変数を見る
SUB VIEW_ENV(P_EnvironmentProperty)
	Dim WSHShell2,WSHEnv2,strList,strEnv
	Set WSHShell2 = WScript.CreateObject("WScript.Shell")
	Set WSHEnv2 = WshShell2.Environment(P_EnvironmentProperty)
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


