' これはUTF-8 bom付きのファイルです
' bomなしだと実行できない（ときがある）
Option Explicit

msgbox "読み込んだ時点で通る", vbOkOnly, "ここ、通る？"

Sub RunTestUtf8
	MsgBox "てすとです", vbInformation, "日本語１２３"
End Sub

