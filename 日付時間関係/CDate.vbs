Option Explicit

Dim rtnCDate

rtnCDate = CDate("2001/4/17")

'IsDate:日付型に変換できるかチェックする関数
MsgBox IsDate(rtnCDate)

MsgBox WeekDay(rtnCDate)
