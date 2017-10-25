Option Explicit

Dim i
For i = 0 To WScript.Arguments.Count-1
	InputBox "落としたファイル～☆","コピーしてね～！",WScript.Arguments(i)
Next

