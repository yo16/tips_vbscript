Option Explicit


Dim testStr
testStr = "abcd123efg"



Dim startPos
startPos = InStr(testStr, "123")

Dim endPos
endPos = InStr(testStr, "f")


Dim leftStr
leftStr = Left(testStr, startPos-1)
'msgbox leftStr

Dim rightStr
rightStr = Right(testStr, Len(testStr)-endPos+1)
'msgbox rightStr


Dim repStr
repStr = leftStr & "999" & rightStr
msgbox repStr
