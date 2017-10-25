Option Explicit

Dim i
For i = 0 To WScript.Arguments.Count-1
	MsgBox WScript.Arguments(i)
Next

