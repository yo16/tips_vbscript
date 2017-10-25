' 自分のファイル名を取得

' ファイル名だけ
MsgBox WScript.ScriptName

' ファイル名を含むフルパス
MsgBox WScript.ScriptFullName

' フォルダ名だけ（\を含めない←マイナス1の分）
MsgBox Left( _
	WScript.ScriptFullName, _
	Len(WScript.ScriptFullName) - Len(WScript.ScriptName) - 1 _
)
