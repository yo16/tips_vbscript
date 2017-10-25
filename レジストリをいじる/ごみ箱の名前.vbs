Option Explicit


'**	カスタマイズしよう！シリーズ
'**		ごみ箱の名前を変える編
'**				2001/01/29



Dim objWshShell
Set objWshShell = WScript.CreateObject("WScript.Shell")

Dim regPath
regPath = "HKCR\CLSID\{645FF040-5081-101B-9F08-00AA002F954E}\"

Dim oldName
oldName = objWshShell.RegRead(regPath)

Dim newName
newName = InputBox("ごみ箱に" & vbCrLf & "新しい名前を" & vbCrLf & "つけてあげましょう♪","（~▽~＠）♪♪♪",oldName)

If ( (newName = "") or (newName = oldName) ) Then
	MsgBox "変えませんでしたとさ。。",0,"(×_×; )"
Else
	MsgBox "「" & newName & "」に変えときました。" & vbCrLf & "再ログオン時から有効です。",0,"(￣ー￣)v"
	MsgBox "言い忘れましたが、",0,"ちょっとオシラセ。(・_・?)"
	MsgBox "レジストリをいじってます。" & vbCrLf & "なんかあっても責任は負いません。" & vbCrLf & "ごめんね～。しょうがないよね～。",16,"「(≧ロ≦) アイタッ！"
	Dim modori
	modori = MsgBox("元の「" & oldName & "」に戻しますか？",292,"実はまだ間に合う。ヽ（∇⌒ヽ）（ノ⌒∇）ノ")

	If (modori = 7) Then
		objWshShell.RegWrite regPath,newName
		MsgBox "マジで変えました。" & vbCrLf & "戻したいときはもう一度やってね～♪",0,"(o･･o)/~マタネェ～"
	Else
		MsgBox "いくじなし。",,"o(・_・)9 アッパー！"
	End If

End If



