Option Explicit

' Function呼び出し

Dim returnValue
returnValue = func_A(1, 2)




Function func_A(paramA, paramB)
	' 戻り値設定
	func_A = paramA+paramB
	
	
	Exit Function		' ←これが必要
	
	msgbox "ココ通る？"
	
End Function
