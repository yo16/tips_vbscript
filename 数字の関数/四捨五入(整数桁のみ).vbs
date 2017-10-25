Option Explicit

msgbox ShisyaGonyu(3.14,-1)

Function ShisyaGonyu(pNumber,pKeta)
	Dim numA,numB,numC,numD,numE,numF,numRtn
	numA = pNumber \ (10^pKeta)
	numB = numA * (10^pKeta)
	numC = pNumber - numB
	numF = 10^(pKeta-1)
	numD = numC - (( numC \ numF ) * numF)
	numE = (numC - numD) \ numF
	If (numE >= 5) Then
		numRtn = numB + (10^(pKeta))
	Else
		numRtn = numB
	End If
	ShisyaGonyu = numRtn
End Function


