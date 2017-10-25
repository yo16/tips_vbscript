' Mid関数

' 引数２の文字から、引数３の数分抜き出す
' 引数３は、終わりの位置ではなく、長さであることに注意
MsgBox Mid("abcdefg", 2, 3)
' bcd


' 引数３を省略すると最後まで
MsgBox Mid("abcdefg", 2)
' bcdefg

' 参考：Rightは、右から文字数分抜き出す
MsgBox Right("abcdefg", 3)
' efg
