Option Explicit

Dim motoStr
motoStr = "param=12345=abc"
'          123456789012345

Dim rtnArray
rtnArray = Split(motoStr, "=")

msgbox rtnArray(0)
msgbox rtnArray(1)
msgbox rtnArray(2)

