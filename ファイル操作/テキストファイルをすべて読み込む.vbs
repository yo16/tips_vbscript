Option Explicit

Dim objFS, objTS
Dim strLine, strTemp

Set objFS = WScript.CreateObject("Scripting.FileSystemObject")

Set objTS = objFS.OpenTextFile("Sample.txt",1)
strLine = ""

Do Until objTS.AtEndOfStream
	strTemp = objTS.ReadLine
	If Not(strTemp = "") Then
		WScript.Echo strTemp
	End If
Loop
objTS.Close

Set objTS = Nothing
Set objFS = Nothing
