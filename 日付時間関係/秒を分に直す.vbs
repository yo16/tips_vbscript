Option Explicit

Dim byou
byou = 150

msgbox ByouToFun(byou)


'60�Ŋ����Đ؂�̂Ă邾������H
Function ByouToFun(pTime)
	Dim rtnByou
	rtnByou = Int(pTime)

	ByouToFun = rtnByou \ 60
End Function


