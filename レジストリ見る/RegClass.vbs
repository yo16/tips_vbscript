Option Explicit
' レジストリエクスポートファイルのクラス
' 
' ClsRegKey
'		AddSubKey( pRegKey )
'		AddValue( pValue )
'		GetSubKeysCount()
'		GetValuesCount()
'		GetSubKeyObjAt( pIndex )
'		GetValueObjAt( pIndex )
'		GetSubKeyObjByKey( pKey )
'		GetValueByName( pName )
'		GetKey()
'		SetKey( pKey )
' ClsRegValue
'		GetName()
'		GetValue()
'		SetName( pName)
'		SetValue( pValue )
'
' 2006/10/11 ikeda


Class ClsRegKey
	'********************************************************************
	' メンバ変数(Private)
	'********************************************************************
	Dim Key			' As String					キー値
	Dim SubKeys		' As Array(ClsRegKey)		サブキー
	Dim Values		' As Array(ClsRegValue)		値
	
	'********************************************************************
	' メンバ関数
	'********************************************************************
	'--------------------------------------------------------------------
	' コンストラクタ
	'--------------------------------------------------------------------
	Private Sub Class_Initialize
		' ClsRegKeyの配列
		SubKeys = Array()
		' ClsRegValueの配列
		Values = Array()
	End Sub
	
	'--------------------------------------------------------------------
	' デストラクタ
	'--------------------------------------------------------------------
	Private Sub Class_Terminate
		Set SubKeys = Nothing
		Set Values = Nothing
	End Sub
	
	'--------------------------------------------------------------------
	' AddSubKey
	' サブキーを１つ増やす
	'--------------------------------------------------------------------
	Public Sub AddSubKey( pRegKey )	' pRegKey As ClsRegKey
		' 再定義
		ReDim Preserve SubKeys(UBound(SubKeys)+1)
		' 追加
		Set SubKeys(UBound(SubKeys)) = pRegKey
	End Sub
	
	'--------------------------------------------------------------------
	' AddValue
	' 値を１つ増やす
	'--------------------------------------------------------------------
	Public Sub AddValue( pValue )	' pRegKey As ClsRegValue
		' 再定義
		ReDim Preserve Values(UBound(Values)+1)
		' 追加
		Set Values(UBound(Values)) = pValue
	End Sub
	
	'--------------------------------------------------------------------
	' GetSubKeysCount
	' サブキーの数を返す
	'--------------------------------------------------------------------
	Public Function GetSubKeysCount()
		GetSubKeysCount = UBound(SubKeys) + 1
	End Function
	
	'--------------------------------------------------------------------
	' GetValuesCount
	' 値の数を返す
	'--------------------------------------------------------------------
	Public Function GetValuesCount()
		GetValuesCount = UBound(Values) + 1
	End Function
	
	'--------------------------------------------------------------------
	' GetSubKeyObjAt
	' サブキーを１つ返す
	'--------------------------------------------------------------------
	Public Function GetSubKeyObjAt( pIndex )	' pIndex As Integer
		Set GetSubKeyObjAt = SubKeys( pIndex )
	End Function
	
	'--------------------------------------------------------------------
	' GetValueAt
	' 値オブジェクトを１つ返す
	'--------------------------------------------------------------------
	Public Function GetValueObjAt( pIndex )	' pIndex As Integer
		Set GetValueObjAt = Values( pIndex )
	End Function
	
	'--------------------------------------------------------------------
	' GetSubKeyObjByKey
	' Keyをキーにしてサブキーオブジェクトを検索し、そのオブジェクトを返す
	'--------------------------------------------------------------------
	Public Function GetSubKeyObjByKey( pKey )	' pKey As String
		Dim rtnIndex
		
		Dim i, skCount
		skCount = GetSubKeysCount()
		For i = 0 to skCount-1
			If ( SubKeys(i).GetKey() = pKey ) Then
				rtnIndex = i
			End If
		Next
		
		Set GetSubKeyObjByKey = SubKeys( rtnIndex )
	End Function
	
	'--------------------------------------------------------------------
	' GetValueByName
	' 名前をキーにして値オブジェクトを検索し、その値を返す
	'--------------------------------------------------------------------
	Public Function GetValueByName( pName )	' pName As String
		Dim rtnStr
		rtnStr = ""
		
		Dim i, vCount
		vCount = GetValuesCount()
		For i = 0 to vCount-1
			If ( Values(i).GetName() = pName ) Then
				rtnStr = Values(i).GetValue()
			End If
		Next
		
		GetValueByName = rtnStr
	End Function
	
	
	'--------------------------------------------------------------------
	' getter
	'--------------------------------------------------------------------
	Public Function GetKey()
		GetKey = Key
	End Function
	
	'--------------------------------------------------------------------
	' setter
	'--------------------------------------------------------------------
	Public Sub SetKey( pKey )	' pKey As String
		Key = pKey
	End Sub
	
End Class

Class ClsRegValue
	'********************************************************************
	' メンバ変数(Private)
	'********************************************************************
	Dim Name
	Dim Value
	
	'********************************************************************
	' メンバ関数
	'********************************************************************
	'--------------------------------------------------------------------
	' getter
	'--------------------------------------------------------------------
	Public Function GetName()
		GetName = Name
	End Function
	Public Function GetValue()
		GetValue = Value
	End Function
	
	'--------------------------------------------------------------------
	' setter
	'--------------------------------------------------------------------
	Public Sub SetName( pName )
		Name = pName
	End Sub
	Public Sub SetValue( pValue )
		Value = pValue
	End Sub
	
End Class
