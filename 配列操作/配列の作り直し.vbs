'��蒼������O�̂͏������Ⴄ�񂾂�

Option Explicit


Dim testStr
Dim tmpArray

testStr = "a,b,c,d,e"
tmpArray = Split(testStr,",")
MsgBox "�z��̌�=>"&UBound(tmpArray)

testStr = "a,b,c"
tmpArray = Split(testStr,",")
MsgBox "�z��̌�=>"&UBound(tmpArray)


