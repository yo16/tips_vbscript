Option Explicit

Dim strA, strB ,pos

strA = "=abc=de=fgh"
'       1234567890


strB = "="


' �����J�n�ʒu���w��
'pos = Instr( 0, strA, strB )		' �O�������Ǝ��s�G���[�ɂȂ�
'msgbox "ans:"+CStr(pos)

pos = Instr( 1, strA, strB )	' �擪�͂P
msgbox "ans:"+CStr(pos)
' ans:1

pos = Instr( 8, strA, strB )
msgbox "ans:"+CStr(pos)
' ans:8

pos = Instr( 9, strA, strB )
msgbox "ans:"+CStr(pos)
' ans:0

pos = Instr( 100, strA, strB )		' ��낷���͖��Ȃ��A0���Ԃ�
msgbox "ans:"+CStr(pos)
' ans:0


' ���݂��Ȃ��ꍇ����L9�Ɠ���
pos = InStr( 1, "abc", "=")
msgbox "ans:"+CStr(pos)
' ans:0
