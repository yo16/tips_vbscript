Option Explicit
'DateDiff
' 2015/7/30

Dim dtToday
dtToday = CDate("2015/7/30")
Dim dtOneday
dtOneday = CDate("2015/5/1")

' DateDiff�i��R�����|��Q�����j
MsgBox DateDiff("d", dtOneday, dtToday)
' �� 90
MsgBox DateDiff("d", dtToday, dtOneday)
' �� -90
' ��R�����|��Q�����̒l��Ԃ�
' �P�ڂ̈���
' �ݒ�l	���e 
' yyyy		�N 
' q			�l���� 
' m			�� 
' y			�N�ԒʎZ�� 
' d			�� 
' w			�T�� 
' ww		�T 
' h			�� 
' n			�� 
' s			�b 
