Option Explicit


' [ 型 ]
' InputBox(prompt[, title][, default][, xpos][, ypos][, helpfile, context])
'
' prompt	: 入力を促すメッセージ
' title		: ウィンドウのタイトル
' default	: TextFieldにあらかじめ入れる文字列
' xpos		: 画面の左端からの距離(twip 単位)
' ypos		: 画面の上端からの距離(twip 単位)
' helpfile	: ヘルプファイルをつけることができるらしい
' context	: ヘルプファイルの引数らしい


Dim modori

' 基本的にこれくらいわかってればいいんじゃん？
modori = InputBox("prompt", "title", "default")


' カクニン
msgbox modori
