Option Explicit

Dim strA, strB ,pos

strA = "abCdefCg"
strB = "C"

pos = InStr(strA, strB )
msgbox "�ʏ�:"+CStr(pos)	' 3

pos = InStrRev(strA, strB )
msgbox "�t:"+CStr(pos)		' 7



' �����T���v��
' �Ō��C���O���擾
Dim targetStr
targetStr = Left(strA,pos-1)
msgbox targetStr			' abCdef
