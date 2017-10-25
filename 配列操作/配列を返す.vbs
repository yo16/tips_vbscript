Option Explicit

' 配列を返す
' 2015/7/29

Dim ret
ret = MakeArray()

Dim str
str = ""
Dim i
For i=0 to UBound(ret)
	str = str & ret(i) & vbCrLf
Next
MsgBox str


Function MakeArray()
	Dim ary
	ary = Array("a", "b", "c")
	MakeArray = ary
End Function

