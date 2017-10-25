Option Explicit

Dim fileName
fileName = "abc.txt"


Dim objFS,objTS
Set objFS = WScript.CreateObject("Scripting.FileSystemObject")

If objFS.FileExists(fileName) Then
	Dim objFile
	Set objFile = objFS.GetFile(fileName)
	Set objTS = objFile.OpenAsTextStream(8)
	objTS.WriteLine "開きましたよ。" & Now
Else
	Set objTS = objFS.CreateTextFile(fileName)
	objTS.WriteLine "開きましたよ。" & Now
end If
objTS.Close


