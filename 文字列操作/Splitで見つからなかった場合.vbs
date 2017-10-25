Option Explicit
' Split‚ÅŒ©‚Â‚©‚ç‚È‚©‚Á‚½ê‡A(0)‚É“ü‚é‚Ì‚©H

Dim aryFound
aryFound = Split("abcde", "/")

MsgBox aryFound(0)		' abcde
MsgBox UBound(aryFound)	' 0
' (0)‚É‚Í‚¢‚èAUBound‚Í0‚É‚È‚é
