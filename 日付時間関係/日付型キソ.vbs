' “ú•tŒ^‚ÌƒLƒ\

Dim v_dateA, v_dateB
Dim v_diffAB



' •¶š—ñ‚ğ“ú•tŒ^‚É•ÏŠ·
v_dateA = CDate("2004/4/1")
v_dateB = CDate("2004/5/1")

msgbox v_dateA
msgbox v_dateB


' ‚¢‚ë‚¢‚ë‚È‚â‚è•û‚ÅA‚Q‚Â‚Ì·‚ğæ‚Á‚Ä‚İ‚é
v_diffAB = DateDiff("y", v_dateA, v_dateB)		' ”NŠÔ’ÊZ“ú ‚`‚a
msgbox v_diffAB
v_diffAB = DateDiff("d", v_dateA, v_dateB)		' “ú•t ‚`‚a
msgbox v_diffAB
v_diffAB = DateDiff("y", v_dateB, v_dateA)		' ”NŠÔ’ÊZ“ú ‚a‚`
msgbox v_diffAB
v_diffAB = DateDiff("d", v_dateB, v_dateA)		' “ú•t ‚a‚`
msgbox v_diffAB

' Œ‹‰ÊË ‡”Ô‚ÍŠÖŒW‚ ‚éB‘æ‚Rˆø”|‘æ‚Qˆø”‚ğ•Ô‚·I


