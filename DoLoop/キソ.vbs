Option Explicit

Dim idx,str

idx = 0
str = ""
Do While (idx<10)
	str = str & idx
	idx = idx + 1
Loop
msgbox "�P���"&str

idx = 0
str = ""
Do
	str = str & idx
	idx = idx + 1
Loop While (idx<10)
msgbox "�Q���"&str
