Option Explicit


Dim testStr
testStr = "a,b,c,d,e,f,g"

MsgBox getCsvStr(testStr,2)



'****************************************
'csvStr��partNumber�Ԗڂ̕������Ԃ��֐�
'****************************************
Function getCsvStr(csvStr,partNumber)
	Dim csvStrArray
	csvStrArray = Split(csvStr,",")
	getCsvStr = csvStrArray(partNumber-1)
End Function


