'ì‚è’¼‚µ‚½‚ç‘O‚Ì‚ÍÁ‚¦‚¿‚á‚¤‚ñ‚¾‚Ë

Option Explicit


Dim testStr
Dim tmpArray

testStr = "a,b,c,d,e"
tmpArray = Split(testStr,",")
MsgBox "”z—ñ‚ÌŒÂ”=>"&UBound(tmpArray)

testStr = "a,b,c"
tmpArray = Split(testStr,",")
MsgBox "”z—ñ‚ÌŒÂ”=>"&UBound(tmpArray)


