Option Explicit

Dim idx
idx = 1
Dim str
str = ""
Do While (idx < 10)
	str = str & idx
	
	' 途中で終了
	If ( idx = 5 ) Then
		Exit Do
		' Exit は、
		' Do...Loop ループ、For...Next ループ、Function プロシージャまたは Sub プロシージャから抜け出すためのフロー制御ステートメントです。

		
	End If
	
	
	idx = idx + 1
Loop

msgbox str

