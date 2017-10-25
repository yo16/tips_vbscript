Option Explicit

Dim strA, strB ,pos

strA = "abcdefg"


strB = "h"
pos = Instr(strA, strB )
msgbox "ans:"+CStr(pos)
' 0


strB = "c"
pos = Instr(strA, strB )
msgbox "ans:"+CStr(pos)
' 3
msgbox Left(strA, pos)
' abc

strB = "a"
pos = Instr(strA, strB )
msgbox "ans:"+CStr(pos)
' 1



' ‰üs‚àŒŸõ‚Å‚«‚é‚©‚È¨‚Å‚«‚é
Dim strC
strC = "abc" & vbCrLf & "def"
pos = Instr(strC, vbCrLf)
' ‰üs‚Ì‘O‚Ü‚Å
msgbox "ans:" & Left(strC, pos)
' abc
