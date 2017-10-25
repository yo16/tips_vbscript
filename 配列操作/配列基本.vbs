Option Explicit

'初期設定
Dim array0
array0 = Array()
' 終わりのとき
Set array0 = Nothing



'使用方法１
Dim array1
array1 = Array("a","b","c")		'Array関数を使用

'配列はゼロベースで()を使う
msgbox array1(0) & "-" & array1(1) & "-" & array1(2)



'使用方法２
Dim array2(2)		'個数-1を宣言（ゼロの分）
array2(0) = "A"
array2(1) = "B"
array2(2) = "C"

'配列はゼロベースで()を使う
msgbox array2(0) & "-" & array2(1) & "-" & array2(2)

' 要素数
MsgBox "UBound(array2)=" & UBound(array2)

'ダメな例１
'Dim array3			'気持ちは配列(実際は０次元)
'array3(0) = "a"		'配列に入れてみる

' だめな例２
'Dim array4
'array4 = Array(3)		' 要素ができてない。どういう状態か不明。

' だめな例３
Dim d1
d1 = 3
'Dim array5(d1)		' 要素数の定義に変数を使えない
Dim array5
ReDim array5(d1)	' →回避方法：いったん型なしで定義した後、ReDimで配列にする
array5(0) = "x"


