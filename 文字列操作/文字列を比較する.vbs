Option Explicit

'Const vbBinaryCompare = 0
'Const vbTextCompare = 1

Dim strA
Dim strB
Dim strC
Dim strD

strA = "abc"
strB = "ABC"
strC = "AbC"
strD = ""

MsgBox "vbBinaryCompare:0" & vbCrLf & "vbTextCompare:1"

MsgBox "StrComp(""" & strA & """,""" & strB & """,vbTextCompare) => "&StrComp(strA,strB,vbTextCompare)
MsgBox "StrComp(""" & strA & """,""" & strC & """,vbTextCompare) => "&StrComp(strA,strC,vbTextCompare)
MsgBox "StrComp(""" & strB & """,""" & strC & """,vbTextCompare) => "&StrComp(strB,strC,vbTextCompare)

MsgBox """" & strA & """ = """ & strB & """ : " & (strA = strB)
MsgBox """" & strA & """ = """ & strC & """ : " & (strA = strC)
MsgBox """" & strB & """ = """ & strC & """ : " & (strB = strC)
MsgBox """" & strA & """ = """ & strA & """ : " & (strA = strA)





