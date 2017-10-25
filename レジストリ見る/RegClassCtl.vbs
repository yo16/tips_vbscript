Option Explicit
' GetRegClass
' ClsRegKey�֒l��ݒ肵���I�u�W�F�N�g��Ԃ��֐�
'
' �߂�l	ClsRegKey�N���X
'
' 2006/10/11 ikeda

Function GetRegClass(fileName, key)
	' �߂�l
	Dim rtnObj
	Set rtnObj = new ClsRegKey
	
	' key��ݒ�
	rtnObj.SetKey(key)
	
	
	Dim objFS
	Set objFS = WScript.CreateObject("Scripting.FileSystemObject")
	Dim objTS
	Set objTS = objFS.OpenTextFile(fileName, 1, false, true)
	
	Dim tmpLine
	tmpLine = objTS.ReadLine	' �P�s�ڂ͎̂Ă�
	
	Dim curKeyStr, tmpKeyStr, subKeyStr
	Dim isTargetSubKey
	isTargetSubKey = false
	Dim posInStrRev, objSubKey, objVal
	Do Until objTS.AtEndOfStream
		tmpLine = objTS.ReadLine
		If (tmpLine = "") Then
			' �L�[
			If ( Not objTS.AtEndOfStream ) Then
				tmpLine = objTS.ReadLine
				
				' �J�����g�̃L�[
				curKeyStr = Mid( tmpLine, 2, Len(tmpLine)-2 )
				
				' �p�����[�^�̃T�u�L�[���H
				' �Q������([�̌�)����A��납�猟������\�̑O�܂ł��Akey�Ɣ�r����
				posInStrRev = InStrRev( tmpLine, "\" )
				If ( posInStrRev > 1 ) Then
					tmpKeyStr = Mid( tmpLine , 2, posInStrRev-2 )
					If ( key = tmpKeyStr ) Then
						isTargetSubKey = true
					End If
				End If
				
				If ( isTargetSubKey ) Then
					' �ċA�I�ɃT�u�L�[�̃I�u�W�F�N�g���P�ǉ�
					subKeyStr = Mid(tmpLine, 2, Len(tmpLine)-2)
					Set objSubKey = GetRegClass( fileName, subKeyStr )
					rtnObj.AddSubKey( objSubKey )
				End If
			End If
		Else
			' �l
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
