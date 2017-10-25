Option Explicit

Dim myArray
myArray = Array("A", "B", "C")

MsgBox isExists(myArray, "A")
MsgBox isExists(myArray, "B")
MsgBox isExists(myArray, "C")
MsgBox isExists(myArray, "a")
MsgBox isExists(myArray, "X")


' =================================================
' isExists
' ^^^^^^^^
' 配列にvalueが存在するかチェックする
' 戻り値 : [ TRUE:存在する | FALSE:存在しない ]
' =================================================
Function isExists(aryCheck, value)
	' 戻り値
	Dim returnValue
	returnValue = False
	
	' ループで使用する変数
	Dim i, intMaxValue
	i = 0
	intMaxValue = Ubound(aryCheck)
	For i = 0 to intMaxValue
		if ( aryCheck(i) = value ) Then
			' 同じだったらTrueを入れる
			returnValue = True
		End If
	Next
	
	isExists = returnValue
	
End Function
