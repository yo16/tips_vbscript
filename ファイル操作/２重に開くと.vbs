' ÇQèdÇ…äJÇ≠Ç∆Ç«Ç§Ç»ÇÈÇ©
' Å®Ç§Ç‹Ç≠Ç‚ÇÈ
' 2006/10/11 ikeda


Dim objFS, objTS
Dim strLine

Set objFS = WScript.CreateObject("Scripting.FileSystemObject")
Set objTS = objFS.OpenTextFile("Sample.txt",1)
strLine = objTS.ReadLine
msgbox strLine

open2()

strLine = objTS.ReadLine
msgbox strLine
strLine = objTS.ReadLine
msgbox strLine


objTS.Close


Sub open2()
	msgbox "open2"
	
	Dim objFS, objTS
	Dim strLine
	
	Set objFS = WScript.CreateObject("Scripting.FileSystemObject")
	Set objTS = objFS.OpenTextFile("Sample.txt",1)
	strLine = objTS.ReadLine
	msgbox strLine
	strLine = objTS.ReadLine
	msgbox strLine
	strLine = objTS.ReadLine
	msgbox strLine
	
	msgbox "open2 end"
	
	objTS.Close
	
End Sub
