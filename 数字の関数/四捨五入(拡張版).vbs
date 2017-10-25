Option Explicit

Dim clcNumber,ichi
clcNumber = InputBox ("四捨五入したい数値を入力!")
ichi = InputBox ("四捨五入したい桁を入力！(小数はマイナス)")

msgbox shisyaGonyu(clcNumber,ichi),,"答え！"

Function shisyaGonyu(pNumber,pKeta)
	If (pKeta = 0) Then
		shisyaGonyu = pNumber
		Exit Function
	End If

	Dim decKeta
	If (pKeta > 0) Then
		decKeta = 10 ^ pKeta
	Else
		decKeta = 10 ^ (pKeta + 1)
	End If

	Dim tmpNumber1
	tmpNumber1 = pNumber / decKeta

	Dim tmpNumber2
	tmpNumber2 = Fix(tmpNumber1)

	Dim tmpNumber3
	tmpNumber3 = (tmpNumber1 - tmpNumber2)*10

	Dim tmpNumber4
	If (tmpNumber3 >= 5) Then
		tmpNumber4 = (tmpNumber2 + 1) * decKeta
	Else
		tmpNumber4 = tmpNumber2 * decKeta
	End If

	shisyaGonyu = tmpNumber4
End Function
