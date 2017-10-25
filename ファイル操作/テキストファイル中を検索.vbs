Option Explicit


Dim objFS
Set objFS = WScript.CreateObject("Scripting.FileSystemObject")

Dim objTS
Set objTS = objFS.OpenTextFile("sample.txt",1)

Dim fileStr
fileStr = objTS.ReadAll

objTS.Close




' åüçı
Dim pos
pos = Instr(fileStr,"CreateObject")

msgbox pos & "ï∂éöñ⁄"
