'*************************************
'substr�n�̊֐�
'
' Left�AMid�ARight������
'*************************************

Option Explicit

Dim testStr
testStr = "abcdefg"

MsgBox "Left:" & Left(testStr, 3)
' abc
MsgBox "Mid:" & Mid(testStr, 2, 3)
' bcd
MsgBox "Right:" & Right(testStr, 3)		' �J�n�ʒu����E
' efg



'������́A���P�������
MsgBox "���P�������F" & Right(testStr, Len(testStr)-1)
' bcdefg
