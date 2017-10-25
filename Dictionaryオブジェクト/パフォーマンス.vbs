Option Explicit

' ハッシュのパフォーマンス測定
' 2016/10/20 y.ikeda
' 登録した数に対して、検索速度がどの程度遅くなるか


' テスト検索数
Dim testSearchNumber
testSearchNumber = 1000000


' 登録数を変えて、何度か計測
TestUseDictionary 100000
' → 7s
TestUseDictionary 200000
' → 12s
TestUseDictionary 300000
' → 17s
TestUseDictionary 400000
' → 22s

' そんなに悪くない印象
' 極端に悪化することもない





Sub TestUseDictionary( registNumber )
	Dim objDic
	Dim i
	Dim startDt, endDt
	Dim spentTime_s
	Dim testKeyNo
	Dim testItem
	Set objDic = CreateObject("Scripting.Dictionary")
	For i=0 To registNumber
		objDic.Add "KEY-"&i, "ITEM-"&i
	Next

	Randomize

	' 計測開始
	startDt = Now

	For i=0 To testSearchNumber
		' 0～registNumber-1のランダムな値を設定
		testKeyNo = Int( registNumber*Rnd )
		testItem = objDic.Item( "KEY-"&testKeyNo )
	Next

	' 計測終了
	endDt = Now
	spentTime_s = DateDiff("s", startDt, endDt)

	MsgBox "登録数:" & registNumber & vbCrLf & _
		spentTime_s & "(s)"

End Sub
