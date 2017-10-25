Option Explicit

' •¶š—ñ‚Ì“ú‚ğDateTimeŒ^‚Ö•ÏŠ·
' 2015/7/30

Dim dt
dt = CDate("2015/7/30 13:4:6")

' DateTimeŒ^‚Æ‚µ‚Ä³‚µ‚­“ü‚Á‚Ä‚¢‚é‚©Šm”F
MsgBox FormatDateTime(dt, 0)
' ¨ 2015/07/30 13:04:06
MsgBox Year(dt)
' ¨ 2015

