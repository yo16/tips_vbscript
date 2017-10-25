Option Explicit


Dim array1
array1 = Array("a","b","c")
' この後ろに要素をもう１個付け加えたいなぁ。。

'*****ここ。*****
' 配列を長く再定義(値を保持したいときはPreserveを指定する)
ReDim Preserve array1(UBound(array1)+1)
' 配列の最後に要素を追加
array1(UBound(array1)) = "x"

msgbox array1(0)
msgbox array1(1)
msgbox array1(2)
msgbox array1(3)

