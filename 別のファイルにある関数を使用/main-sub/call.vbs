Dim WSHShell,intErrCode
Set WSHShell = WScript.CreateObject("WScript.Shell")
intErrCode=WSHShell.Run("D:\wsh\main\main.wsf",1,True)

Select Case intErrCode
	Case -1 MsgBox "ダイアログは自動的に閉じられました。"
	Case -2 MsgBox	"エラーなんだな！"
	Case vbYes MsgBox "「はい」を押しました。:  " & vbYes
	Case vbNo MsgBox "「いいえ」を押しました。:  " & vbNo
	Case vbCancel MsgBox "「キャンセル」を押しました。:  " & vbCancel
End Select

