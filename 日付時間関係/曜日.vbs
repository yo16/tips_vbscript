Option Explicit

Dim today
today = Date()

msgbox FormatDateTime(today,2) & "��" & Weekday(today) & "�j��", vbOkOnly, "1 (���j) �` 7 (�y�j) "
' 1 (���j) �` 7 (�y�j) 



Dim yesterday
yesterday = DateAdd("d", -1, today)

msgbox FormatDateTime(yesterday,2) & "��" & Weekday(yesterday) & "�j��", vbOkOnly, "1 (���j) �` 7 (�y�j) "
