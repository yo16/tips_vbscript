Option Explicit

' 変数名に_を使っていいのか疑問をもった。
' けど使ってもよさそう。 2001/11/29


Dim a_a
Dim b_b

a_a = 1
b_b = a_a	_
		+ 1
' _の次は次の行を見るよという意味
' ただし、_の後ろにはコメントを書けない！
a_a = b_b + 1


msgbox a_a
msgbox b_b


