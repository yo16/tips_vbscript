Option Explicit

' 置換
' 2006/10/11 ikeda


msgbox Replace("XXpXXPXXp", "p", "Y")   ' 文字列の最初から、バイナリ モードで比較を行います。"XXYXXPXXY"を返します。
msgbox Replace("XXpXXPXXp", "p", "Y", 3, -1, 1)   ' 3 番目の位置からテキスト モードで比較を行います。"YXXYXXY" を返します。

' ヒットしないときは何もしない
MsgBox Replace("abcde", "x", "y")
