Option Explicit


Dim objFS
Set objFS = WScript.CreateObject("Scripting.FileSystemObject")

If (objFS.FileExists("�R�s�`abc.txt") <> -1) Then
	objFS.CopyFile "abc.txt","�R�s�`abc.txt"
	WScript.Quit
End If

Dim idx
idx = 1
Do
	If (objFS.FileExists("�R�s�`"&idx&"abc.txt") <> -1) Then
		objFS.CopyFile "abc.txt","�R�s�`"&idx&"abc.txt"
		WScript.Quit
	End If
	idx = idx + 1
Loop
