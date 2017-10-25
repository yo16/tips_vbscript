Option Explicit

Dim readFile1,readFile2
readFile1 = "a.txt"
readFile2 = "b.txt"

Dim writeFile
writeFile = "見比べ結果.txt"


Dim objFS
Set objFS = WScript.CreateObject("Scripting.FileSystemObject")

Dim objTSread1,objTSread2
Set objTSread1 = objFS.OpenTextFile(readFile1,1)
Set objTSread2 = objFS.OpenTextFile(readFile2,1)

Dim objTSwrite
Set objTSwrite = objFS.CreateTextFile(writeFile,1)

Dim fileStr1,fileStr2
Dim lineNumber,errorNumber
lineNumber = 1
errorNumber = 0
Do Until (objTSread1.AtEndOfStream or objTSread2.AtEndOfStream)
	fileStr1 = objTSread1.ReadLine
	fileStr2 = objTSread2.ReadLine
	If (fileStr1 <> fileStr2) Then
		objTSwrite.WriteLine lineNumber&"行目が違います！"
		objTSwrite.WriteLine "1: "&fileStr1
		objTSwrite.WriteLine "2: "&fileStr2
		errorNumber = errorNumber + 1
	End If

	lineNumber = lineNumber + 1
Loop

If Not (objTSread1.AtEndOfStream and objTSread2.AtEndOfStream) Then
	If objTSread1.AtEndOfStream Then
		MsgBox "ファイル"&readFile2&"の方が長いみたいですよ？"
	Else
		MsgBox "ファイル"&readFile1&"の方が長いみたいですよ？"
	End If
End If


objTSread1.Close
objTSread2.Close
objTSwrite.Close


msgbox errorNumber & "/" & lineNumber & "行違いました！",,"終ったよ～♪"


