Option Explicit

Dim strA, strB ,pos

strA = "=abc=de=fgh"
'       1234567890


strB = "="


' 検索開始位置を指定
'pos = Instr( 0, strA, strB )		' 前すぎだと実行エラーになる
'msgbox "ans:"+CStr(pos)

pos = Instr( 1, strA, strB )	' 先頭は１
msgbox "ans:"+CStr(pos)
' ans:1

pos = Instr( 8, strA, strB )
msgbox "ans:"+CStr(pos)
' ans:8

pos = Instr( 9, strA, strB )
msgbox "ans:"+CStr(pos)
' ans:0

pos = Instr( 100, strA, strB )		' 後ろすぎは問題なく、0が返る
msgbox "ans:"+CStr(pos)
' ans:0


' 存在しない場合も上記9と同じ
pos = InStr( 1, "abc", "=")
msgbox "ans:"+CStr(pos)
' ans:0
