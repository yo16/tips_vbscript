Option Explicit
'DateDiff
' 2015/7/30

Dim dtToday
dtToday = CDate("2015/7/30")
Dim dtOneday
dtOneday = CDate("2015/5/1")

' DateDiffi‘æ‚Rˆø”|‘æ‚Qˆø”j
MsgBox DateDiff("d", dtOneday, dtToday)
' ¨ 90
MsgBox DateDiff("d", dtToday, dtOneday)
' ¨ -90
' ‘æ‚Rˆø”|‘æ‚Qˆø”‚Ì’l‚ğ•Ô‚·
' ‚P‚Â–Ú‚Ìˆø”
' İ’è’l	“à—e 
' yyyy		”N 
' q			l”¼Šú 
' m			Œ 
' y			”NŠÔ’ÊZ“ú 
' d			“ú 
' w			T“ú 
' ww		T 
' h			 
' n			•ª 
' s			•b 
