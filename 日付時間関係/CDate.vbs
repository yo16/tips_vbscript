Option Explicit

Dim rtnCDate

rtnCDate = CDate("2001/4/17")

'IsDate:���t�^�ɕϊ��ł��邩�`�F�b�N����֐�
MsgBox IsDate(rtnCDate)

MsgBox WeekDay(rtnCDate)
