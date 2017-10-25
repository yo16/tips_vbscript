Option Explicit

Call sub_A()

msgbox "期待通りにはできなかった。",,"残念無念"

Sub sub_A()
	Dim num1,num2
	
	num1 = 1
	num2 = 2

	msgbox func_add(num1,num2)

'ここには定義できなかった。2001/11/28
'	Function func_add(p1,p2)
'		func_add = p1+p2
'	End Function

End Sub

' 外に出したらOKが出た。
	Function func_add(p1,p2)
		func_add = p1+p2
	End Function
