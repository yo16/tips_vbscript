Option Explicit

Execute ReadFile("ŒÄ‚Î‚ê‚é.vbs")


Call yobaTest()



Function ReadFile(ByVal FileName)
	Const ForReading = 1

	Dim FileShell
	Set FileShell = WScript.CreateObject("Scripting.FileSystemObject")

	ReadFile = FileShell.OpenTextFile(FileName, ForReading, False).ReadAll()
End Function

