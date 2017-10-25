' ループをスキップする
' C言語のcontinue
' VBSには存在しないので、どうするか？
' Ifは、飛ばしたい条件が複数あると、どんどん深くなるためNG
Dim str
str = ""

Dim i
For i=0 to 10
Do
	If (i mod 2 = 0 ) Then Exit Do
	If (i mod 3 = 0 ) Then Exit Do
	str = str & "/" & i
Loop Until 1
Next

MsgBox str
' /1/5/7
' ０～１０で、２の倍数と３の倍数以外
