Option Explicit

msgbox getOneNumber(235,1)

Function getOneNumber(pNumber,pKeta)
	Dim decKeta
	decKeta = 10 ^ (pKeta - 1)

	Dim largeNumber
	largeNumber = Int(pNumber / (decKeta * 10)) * decKeta * 10

	Dim middleNumber
	middleNumber = pNumber - largeNumber

	getOneNumber = Int(middleNumber / decKeta)
End Function
