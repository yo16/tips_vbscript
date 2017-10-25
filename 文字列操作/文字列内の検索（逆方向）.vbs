Option Explicit

Dim strA, strB ,pos

strA = "abCdefCg"
strB = "C"

pos = InStr(strA, strB )
msgbox "’Êí:"+CStr(pos)	' 3

pos = InStrRev(strA, strB )
msgbox "‹t:"+CStr(pos)		' 7



' À‘•ƒTƒ“ƒvƒ‹
' ÅŒã‚ÌC‚æ‚è‘O‚ğæ“¾
Dim targetStr
targetStr = Left(strA,pos-1)
msgbox targetStr			' abCdef
