
Call timePass(3)
msgBox "3ïbå„ÅH"


Sub timePass(passSecond)
	Dim startTime,endTime
	startTime = Timer
	endTime = startTime + passSecond
	Do
	Loop Until (endTime < Timer)
'msgbox startTime&","&endTime
End Sub









