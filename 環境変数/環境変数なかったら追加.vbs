'----------------------------------------------------
' ���ϐ���ݒ肷��
'
' 2006/09/09 ikeda
'----------------------------------------------------
Option Explicit


'----------------------------------------------------
' �ݒ�
'----------------------------------------------------
AddEnv "CLASSPATH", ".", "System"
AddEnv "ENVNAME", "ENVVALUE", "System"







'----------------------------------------------------
' ���C������
'----------------------------------------------------

'----------------------------------------------------
' AddEnv
'
' �w�肳�ꂽ���ϐ���T����
'	�Ȃ���΁A�ϐ��ƒl��ǉ�
'	����΁A�l�Ɏw��l���܂܂�Ă��邩�m�F
'		�Ȃ���΁A�ǉ�
'		����΁A�Ȃɂ����Ȃ�
'----------------------------------------------------
Sub AddEnv(envName, envValue, envObj)
	' WScriptShell���쐬
	Dim WshShell
	Set WshShell = WScript.CreateObject("WScript.Shell")


	' ���ϐ��I�u�W�F�N�g���擾
	Dim envs
	Set envs = WshShell.Environment(envObj)


	' ���ϐ�����CLASSPATH���擾�i�啶���ł��������ł��擾�\�j
	Dim envCurValue
	envCurValue = envs(envName)


	' �������ʂ����āA���ϐ���ݒ肷��
	If ( envCurValue = "" ) Then
		mb "envCurValue is not found"
		' �Ȃ�����
		' �� ���ϐ�[envName]���쐬���A�l��ݒ肷��
		envs.Item(envName) = envValue & ";"
	Else
		mb "envCurValue is found"
		' ������
		' envValue�͓o�^����Ă��邩
		If ( isExists( envCurValue, envValue ) ) Then
			mb envValue & " is found"
			' envValue�͓o�^����Ă���
			' �� �Ȃɂ����Ȃ�
		Else
			mb envValue & " is not found"
			' envValue���o�^����Ă��Ȃ�
			' �� ��ԑO�ɒǉ��o�^����
			envs.Item(envName) = envValue & ";" & envCurValue
		End If
	End If

End Sub

'----------------------------------------------------
' isExists
'
' allPath�ɁAcheckPath���܂܂�Ă��邩���Ȃ������ׂ�
'----------------------------------------------------
Function isExists(allPath, checkPath)
	mb allPath
	Dim returnValue
	returnValue = FALSE
	
	' ";"��allPath����؂��āA���ׂĒ��ׂ�
	Dim continue
	Dim startPos, foundPos
	Dim partPath
	continue = TRUE
	startPos = 1
	Do While ( continue )
		' ;��T��
		foundPos = InStr( startPos, allPath, ";", 1 )
		If ( ( foundPos = Null ) Or ( foundPos = 0 ) ) Then
			' �݂���Ȃ�������A�I��
			continue = FALSE
		Else
			' �݂�������AstartPos�`(foundPos-1)��checkPath�łȂ����`�F�b�N
			partPath = Mid(allPath, startPos, (foundPos-startPos))
			mb partPath
			If ( StrComp( partPath, checkPath ) = 0 ) Then
				' checkPath��������
				returnValue = TRUE
				' �����I��
				continue = FALSE
			ELSE
				' checkPath���Ȃ�����
				' ���̕�������Č���
				startPos = foundPos + 1
			End If
			
		End If
	Loop
	
	' �S���[�v�ŁAcheckPath���Ȃ������ꍇ
	' �Ō��;���疖���܂Ō���
	If ( Not returnValue ) Then
		' startPos�`����������
		partPath = Mid(allPath, startPos, (Len(allPath)-startPos+1))
		mb partPath
		If ( StrComp( partPath, checkPath ) = 0 ) Then
			' checkPath��������
			returnValue = TRUE
		End If
	End If
	
	isExists = returnValue
End Function


' �f�o�b�O�pMsgBox
' �����[�X����FALSE�ɂ���
Sub mb(str)
	If ( FALSE ) Then
		MsgBox str
	End If
End Sub
