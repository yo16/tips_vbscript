Option Explicit

Dim X
Set X = New myDictionary
Dim modori

'Parameter[CpmareMode]を変更するテスト
'	X.CompareMode = 1
'	modori = X.CompareMode
'	MsgBox modori

'Sub[Add]のテスト
	X.Add "p","aaa"

Set X = Nothing


Class myDictionary
	Private dictionaryArray

'* Property [CompareMode](I/O) *
	Private privCompareMode
		'vbBinaryCompare	:0
		'vbTextCpmare		:1
	Public Property Get CompareMode()
		CompareMode = privCompareMode
	End Property
	Public Property Let CompareMode(paramCM)
		If Not ((paramCM = 0) or (paramCM = 1)) Then
			Err.Raise 1,,"パラメータは0か1じゃなくっちゃ！！"
		End If
		privCompareMode = paramCM
	End Property
'strComp

'* Property [Count](O) *
	Private privCount

'* Property [Item](I/O) *
	Private privItem

'* Property [Key](I) *
	Private privKey

'* Sub [Add] *
	Public Sub Add(paramKey,paramItem)
		Dim pArrayCount,newArrayCount
		pArrayCount = UBound(dictionaryArray,1)
msgbox "test pArrayCount is "&pArrayCount
		newArrayCount = pArrayCount + 1
		ReDim Preserve dictionaryArray(newArrayCount,2)'<---ここでエラー落ち。変数はReDimできない？
	End Sub

'* Initialize Terminate *
	Private Sub Class_Initialize
		Dim tmpDictionaryArray(0,2)
		dictionaryArray = tmpDictionaryArray
		privCount = 0
	End Sub
	Private Sub Class_Terminate
		MsgBox("myDictionary が 破棄されました！")
	End Sub
End Class

