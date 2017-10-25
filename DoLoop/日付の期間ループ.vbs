Option Explicit

' “ú•t‚ÌŠúŠÔƒ‹[ƒv


Dim dt1
dt1 = CDate("2015/7/30")
Dim dt2
dt2 = CDate("2015/8/5")


Dim dtCur
dtCur = dt1
Do While( DateDiff("d", dtCur, dt2) >= 0  )
	
	msgbox dtCur
	
	dtCur = DateAdd("d", 1, dtCur)
	
Loop



msgbox "end"