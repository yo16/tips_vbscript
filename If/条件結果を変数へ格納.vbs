Option Explicit

' 条件の結果を変数へ格納
' 関数の戻り値にも使える

Dim a, b, c
a = 1
b = 1
c = 2

Dim ret
ret = (a=b)
msgbox ret
' True

ret = (a=c)
msgbox ret
' False

