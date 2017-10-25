Option Explicit

' 日付の比較を、＜＞で行えるか？

Dim dt1
dt1 = CDate("2016/9/14")
Dim dt2
dt2 = CDate("2016/9/15")

If dt1 < dt2 Then
	MsgBox "dt1 < dt2"
End If
If dt1 = dt2 Then
	MsgBox "dt1 = dt2"
else
	MsgBox "dt1 <> dt2"
End If
