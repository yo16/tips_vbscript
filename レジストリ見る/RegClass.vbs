Option Explicit
' ���W�X�g���G�N�X�|�[�g�t�@�C���̃N���X
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
	' �����o�ϐ�(Private)
	'********************************************************************
	Dim Key			' As String					�L�[�l
	Dim SubKeys		' As Array(ClsRegKey)		�T�u�L�[
	Dim Values		' As Array(ClsRegValue)		�l
	
	'********************************************************************
	' �����o�֐�
	'********************************************************************
	'--------------------------------------------------------------------
	' �R���X�g���N�^
	'--------------------------------------------------------------------
	Private Sub Class_Initialize
		' ClsRegKey�̔z��
		SubKeys = Array()
		' ClsRegValue�̔z��
		Values = Array()
	End Sub
	
	'--------------------------------------------------------------------
	' �f�X�g���N�^
	'--------------------------------------------------------------------
	Private Sub Class_Terminate
		Set SubKeys = Nothing
		Set Values = Nothing
	End Sub
	
	'--------------------------------------------------------------------
	' AddSubKey
	' �T�u�L�[���P���₷
	'--------------------------------------------------------------------
	Public Sub AddSubKey( pRegKey )	' pRegKey As ClsRegKey
		' �Ē�`
		ReDim Preserve SubKeys(UBound(SubKeys)+1)
		' �ǉ�
		Set SubKeys(UBound(SubKeys)) = pRegKey
	End Sub
	
	'--------------------------------------------------------------------
	' AddValue
	' �l���P���₷
	'--------------------------------------------------------------------
	Public Sub AddValue( pValue )	' pRegKey As ClsRegValue
		' �Ē�`
		ReDim Preserve Values(UBound(Values)+1)
		' �ǉ�
		Set Values(UBound(Values)) = pValue
	End Sub
	
	'--------------------------------------------------------------------
	' GetSubKeysCount
	' �T�u�L�[�̐���Ԃ�
	'--------------------------------------------------------------------
	Public Function GetSubKeysCount()
		GetSubKeysCount = UBound(SubKeys) + 1
	End Function
	
	'--------------------------------------------------------------------
	' GetValuesCount
	' �l�̐���Ԃ�
	'--------------------------------------------------------------------
	Public Function GetValuesCount()
		GetValuesCount = UBound(Values) + 1
	End Function
	
	'--------------------------------------------------------------------
	' GetSubKeyObjAt
	' �T�u�L�[���P�Ԃ�
	'--------------------------------------------------------------------
	Public Function GetSubKeyObjAt( pIndex )	' pIndex As Integer
		Set GetSubKeyObjAt = SubKeys( pIndex )
	End Function
	
	'--------------------------------------------------------------------
	' GetValueAt
	' �l�I�u�W�F�N�g���P�Ԃ�
	'--------------------------------------------------------------------
	Public Function GetValueObjAt( pIndex )	' pIndex As Integer
		Set GetValueObjAt = Values( pIndex )
	End Function
	
	'--------------------------------------------------------------------
	' GetSubKeyObjByKey
	' Key���L�[�ɂ��ăT�u�L�[�I�u�W�F�N�g���������A���̃I�u�W�F�N�g��Ԃ�
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
	' ���O���L�[�ɂ��Ēl�I�u�W�F�N�g���������A���̒l��Ԃ�
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
	' �����o�ϐ�(Private)
	'********************************************************************
	Dim Name
	Dim Value
	
	'********************************************************************
	' �����o�֐�
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
