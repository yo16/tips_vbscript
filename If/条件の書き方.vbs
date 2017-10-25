' IF文の条件の書きかた
' （もしかしたらwhile文などにも応用可？）

' 2002/6/13

Option Explicit

If (1) Then
	MsgBox "1:True"		' こっちに入る
Else
	MsgBox "1:False"
End If

If (0) Then
	MsgBox "2:True"
Else
	MsgBox "2:False"		' こっちに入る
End If

If (-1) Then
	MsgBox "3:True"		' こっちに入る
Else
	MsgBox "3:False"
End If


If (True) Then
	MsgBox "4:True"		' こっちに入る
Else
	MsgBox "4:False"
End If


