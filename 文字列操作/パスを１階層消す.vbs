Option Explicit


Dim testStr1 'as String
Dim testStr2 'as String
testStr1 = "C:\900_Programming\VBScript\source\���K�\�[�X\�����񑀍�"

testStr2 = popDirStr( testStr1, 1)

MsgBox(testStr2)




'Function popDirStr( pOldPath as String, pDelDepth as Integer ) As String
Function popDirStr( pOldPath , pDelDepth )
	' �߂�l�̕�����
	Dim rtnStr 'As String
	rtnStr = pOldPath
	
	' �������J�E���g
	Dim delCount 'As Integer
	delCount = 0
	
	
	Dim lastChar 'As String
	
	' �P�������������[�v
	While ( delCount < pDelDepth )
		' �Ō�̕������擾
		lastChar = Right( rtnStr, 1 )
		
		' �Ȃɂ͂Ƃ�����Ō�̕������J�b�g�I
		rtnStr = Left( rtnStr, Len(rtnStr) - 1 )
		
		' �Ō�̕�����\���ǂ������f����
		If ( lastChar = "\" ) Then
			' �Ō�̕������u\�v�̏ꍇ�A�������J�E���g���C���N�������g
			delCount = delCount + 1
		End If
		
	Wend
	
	
	popDirStr = rtnStr
End Function







