Option Explicit

' ������̓�����DateTime�^�֕ϊ�
' 2015/7/30

Dim dt
dt = CDate("2015/7/30 13:4:6")

' DateTime�^�Ƃ��Đ����������Ă��邩�m�F
MsgBox FormatDateTime(dt, 0)
' �� 2015/07/30 13:04:06
MsgBox Year(dt)
' �� 2015

