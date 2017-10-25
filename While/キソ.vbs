Option Explicit

Dim idx
idx = 1
Dim str
str = ""

Do While (idx < 10)
	str = str & idx
	idx = idx + 1
Loop

' «‚±‚ê‚Å‚à‚¢‚¢‚¯‚ÇAExit DoŽg‚¦‚È‚¢‚©‚ç‚È‚é‚×‚­Žg‚¤‚Ì‚æ‚»‚¤‚©‚È
'While (idx < 10)
'	str = str & idx
'	idx = idx + 1
'Wend

msgbox str

