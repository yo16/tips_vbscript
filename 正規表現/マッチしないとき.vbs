Option Explicit

' �}�b�`���Ȃ��Ƃ�
' 2017/11/28 yo16

Dim str
str = "aabbcc"

Dim reg
Set reg = New RegExp
reg.Pattern = "x+"

Dim match
Set match = reg.Execute(str)

msgbox match.Count
' �� 0
