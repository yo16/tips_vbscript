Option Explicit

Dim tmpArray
tmpArray = Array("b","a","c")


Dim arrayValue,strValue
strValue = ""
For Each arrayValue In tmpArray
	strValue = strValue & arrayValue & vbCrLf
Next

MsgBox strValue

