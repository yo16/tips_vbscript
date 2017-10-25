Option Explicit

Dim objFS, objTS
Dim strLine, strTemp

Set objFS = WScript.CreateObject("Scripting.FileSystemObject")

Set objTS = objFS.OpenTextFile("â¸çsÇ™ïœÇ©Ç‡.txt",1)

strLine = objTS.ReadAll

strLine = Replace(strLine, vbCrLf, "<br type=CrLf />")
strLine = Replace(strLine, vbCr, "<br type=Cr />")
strLine = Replace(strLine, vbLf, "<br type=Lf />")

strLine = Replace(strLine, "<br type=CrLf />", "<br type=CrLf />"&vbCrLf)
strLine = Replace(strLine, "<br type=Cr />", "<br type=Cr />"&vbCrLf)
strLine = Replace(strLine, "<br type=Lf />", "<br type=Lf />"&vbCrLf)




objTS.Close


msgbox strLine
