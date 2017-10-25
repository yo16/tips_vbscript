Option Explicit

' 特定の処理を定期的に回す処理
' 2017/2/22 (c) y.ikeda

' 呼び出し間隔（分）
Dim intervalTime : intervalTime = 60

' 動かすVBS
Dim targetScript
targetScript = "一定時間ループ_呼ばれる.vbs"





' 待ち時間（ms）
Dim intTime_ms : intTime_ms = intervalTime * 1000 * 60

' 起動確認
Dim retMsg
retMsg = MsgBox( intervalTime & "分間隔で、起動します。", vbYesNo, "一定時間処理")
If ( retMsg = vbNo ) Then
	WScript.Quit
End If

Dim i
Dim objShell
Set objShell = WScript.CreateObject("WScript.Shell")
For i=0 to 10	' 10回固定
	' 起動
	objShell.Run targetScript, 0, 1
	' ウェイト
	WScript.Sleep intTime_ms
Next

