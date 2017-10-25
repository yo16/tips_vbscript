Option Explicit
' GetRegClass
' ClsRegKeyへ値を設定したオブジェクトを返す関数
'
' 戻り値	ClsRegKeyクラス
'
' 2006/10/11 ikeda

Function GetRegClass(fileName, key)
	' 戻り値
	Dim rtnObj
	Set rtnObj = new ClsRegKey
	
	' keyを設定
	rtnObj.SetKey(key)
	
	
	Dim objFS
	Set objFS = WScript.CreateObject("Scripting.FileSystemObject")
	Dim objTS
	Set objTS = objFS.OpenTextFile(fileName, 1, false, true)
	
	Dim tmpLine
	tmpLine = objTS.ReadLine	' １行目は捨てる
	
	Dim curKeyStr, tmpKeyStr, subKeyStr
	Dim isTargetSubKey
	isTargetSubKey = false
	Dim posInStrRev, objSubKey, objVal
	Do Until objTS.AtEndOfStream
		tmpLine = objTS.ReadLine
		If (tmpLine = "") Then
			' キー
			If ( Not objTS.AtEndOfStream ) Then
				tmpLine = objTS.ReadLine
				
				' カレントのキー
				curKeyStr = Mid( tmpLine, 2, Len(tmpLine)-2 )
				
				' パラメータのサブキーか？
				' ２文字目([の後)から、後ろから検索した\の前までを、keyと比較する
				posInStrRev = InStrRev( tmpLine, "\" )
				If ( posInStrRev > 1 ) Then
					tmpKeyStr = Mid( tmpLine , 2, posInStrRev-2 )
					If ( key = tmpKeyStr ) Then
						isTargetSubKey = true
					End If
				End If
				
				If ( isTargetSubKey ) Then
					' 再帰的にサブキーのオブジェクトを１つ追加
					subKeyStr = Mid(tmpLine, 2, Len(tmpLine)-2)
					Set objSubKey = GetRegClass( fileName, subKeyStr )
					rtnObj.AddSubKey( objSubKey )
				End If
			End If
		Else
			' 値
			If ( curKeyStr = key ) Then
				Set objVal = MakeRegValue( tmpLine )
				rtnObj.AddValue( objVal )
			End If
			
		End If
	Loop
	
	objTS.Close
	Set GetRegClass = rtnObj

End Function

Function MakeRegValue( fileStr )
	Dim rtnObj
	Set rtnObj = new ClsRegValue
	
	Dim splitArray
	splitArray = Split( fileStr, "=" )
	rtnObj.SetName( splitArray(0) )
	If ( UBound(splitArray)>0 ) Then
		rtnObj.SetValue( splitArray(1) )
	End If
	
	Set MakeRegValue = rtnObj
End Function
