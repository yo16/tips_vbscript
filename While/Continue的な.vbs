Option Explicit


Dim idx
idx = 0

Dim sum
sum = 0

Do While (idx < 5)
	If( idx mod 2 = 0 )Then
		idx = idx + 1
		Continue	' ないらしい・・・→If/Elseで。。
	End If
	
	sum = sum + idx
	
	idx = idx + 1
Loop


msgbox sum
