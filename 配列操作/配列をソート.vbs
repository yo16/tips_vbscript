Option Explicit
Dim ary
ary = Array(5, 3, 1, 2, 4)

QuickSort ary, 0, 4

msgbox tostr(ary)


Function tostr(ar)
	Dim str
	str = ""
	Dim i
	For i=0 to UBound(ar)
		str = str & ar(i) & "-"
	Next
	tostr = str
End Function

' http://www.geocities.co.jp/SilkRoad/4511/vb/sample/vbsqsort.htm
'----- QuickSort VBScript Version 1.00 -----
'指定された配列内の要素を、クイックソートによりソートします。
'
'引数 vntArray
'   必ず指定します。ソートを行いたい配列を指定します。例えば要素が
'       vntArray(0) = 123
'       vntArray(1) = 80
'       vntArray(2) = 21
'   の配列を渡した場合、
'       vntArray(0) = 21
'       vntArray(1) = 80
'       vntArray(2) = 123
'   のように正順に整列されます。ただし渡された変数配列を直接比較するため、
'   vntArray が文字列型変数配列とみなされてしまった場合、
'       vntArray(0) = "123"
'       vntArray(1) = "21"
'       vntArray(2) = "80"
'   のように整列されてしまいますので、数値のソートの場合は配列を CLng 関数や
'   CDbl 関数などであらかじめ数値変数配列に変換しておく必要があります。
'
'引数 vntStart
'   必ず指定します。ソートを開始したい要素の番号を指定します。
'
'引数 vntEnd
'   必ず指定します。ソートを終了したい要素の番号を指定します。
'
'再帰的呼び出しを行う関係上、ソート開始・終了番号を省略することは
'できませんのでご注意下さい。
'
Public Sub QuickSort _
    (ByRef vntArray, _
     ByVal vntStart, _
     ByVal vntEnd)

 Dim vntBaseNumber                                      '中央の要素番号を格納する変数
 Dim vntBaseValue                                       '基準値を格納する変数
 Dim vntCounter                                         '格納位置カウンタ
 Dim vntBuffer                                          '値をスワップするための作業域
 Dim i                                                  'ループカウンタ

    If vntStart >= vntEnd Then Exit Sub                 '終了番号が開始番号以下の場合、プロシージャを抜ける
    vntBaseNumber = (vntStart + vntEnd) \ 2             '中央の要素番号を求める
    vntBaseValue = vntArray(vntBaseNumber)              '中央の値を基準値とする
    vntArray(vntBaseNumber) = vntArray(vntStart)        '中央の要素に開始番号の値を格納
    vntCounter = vntStart                               '格納位置カウンタを開始番号と同じにする
    For i = (vntStart + 1) To vntEnd Step 1             '開始番号の次の要素から終了番号までループ
        If vntArray(i) < vntBaseValue Then              '値が基準値より小さい場合
            vntCounter = vntCounter + 1                 '格納位置カウンタをインクリメント
            vntBuffer = vntArray(vntCounter)            'vntArray(i) と vntArray(vntCounter) の値をスワップ
            vntArray(vntCounter) = vntArray(i)
            vntArray(i) = vntBuffer
        End If
    Next
    vntArray(vntStart) = vntArray(vntCounter)           'vntArray(vntCounter) を開始番号の値にする
    vntArray(vntCounter) = vntBaseValue                 '基準値を vntArray(vntCounter) に格納
    Call QuickSort(vntArray, vntStart, vntCounter - 1)  '分割された配列をクイックソート(再帰)
    Call QuickSort(vntArray, vntCounter + 1, vntEnd)    '分割された配列をクイックソート(再帰)

End Sub
