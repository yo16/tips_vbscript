Option Explicit


Dim objFS
Set objFS = WScript.CreateObject("Scripting.FileSystemObject")

If (objFS.FileExists("コピ～abc.txt") <> -1) Then
	objFS.CopyFile "abc.txt","コピ～abc.txt"
	WScript.Quit
End If

Dim idx
idx = 1
Do
	If (objFS.FileExists("コピ～"&idx&"abc.txt") <> -1) Then
		objFS.CopyFile "abc.txt","コピ～"&idx&"abc.txt"
		WScript.Quit
	End If
	idx = idx + 1
Loop
