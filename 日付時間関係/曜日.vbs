Option Explicit

Dim today
today = Date()

msgbox FormatDateTime(today,2) & "‚Í" & Weekday(today) & "—j“ú", vbOkOnly, "1 (“ú—j) ` 7 (“y—j) "
' 1 (“ú—j) ` 7 (“y—j) 



Dim yesterday
yesterday = DateAdd("d", -1, today)

msgbox FormatDateTime(yesterday,2) & "‚Í" & Weekday(yesterday) & "—j“ú", vbOkOnly, "1 (“ú—j) ` 7 (“y—j) "
