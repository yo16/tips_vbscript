

Option Explicit


Function ireru(hikisu)


	If (hikisu = 1) Then '--１が呼んだ
		ireru = "１の戻り値"
		aaa = "111"
	Else
		ireru = "２の戻り値"
		aaa = "222"
	End If


	MsgBox "呼ばれる人の中でちょっとまった。by 呼ぶ人"&hikisu


End Function


