'----------------------------------------------------
' 環境変数を設定する
'
' 2006/09/09 ikeda
'----------------------------------------------------
Option Explicit


'----------------------------------------------------
' 設定
'----------------------------------------------------
AddEnv "CLASSPATH", ".", "System"
AddEnv "ENVNAME", "ENVVALUE", "System"







'----------------------------------------------------
' メイン処理
'----------------------------------------------------

'----------------------------------------------------
' AddEnv
'
' 指定された環境変数を探して
'	なければ、変数と値を追加
'	あれば、値に指定値が含まれているか確認
'		なければ、追加
'		あれば、なにもしない
'----------------------------------------------------
Sub AddEnv(envName, envValue, envObj)
	' WScriptShellを作成
	Dim WshShell
	Set WshShell = WScript.CreateObject("WScript.Shell")


	' 環境変数オブジェクトを取得
	Dim envs
	Set envs = WshShell.Environment(envObj)


	' 環境変数からCLASSPATHを取得（大文字でも小文字でも取得可能）
	Dim envCurValue
	envCurValue = envs(envName)


	' 検索結果を見て、環境変数を設定する
	If ( envCurValue = "" ) Then
		mb "envCurValue is not found"
		' なかった
		' → 環境変数[envName]を作成し、値を設定する
		envs.Item(envName) = envValue & ";"
	Else
		mb "envCurValue is found"
		' あった
		' envValueは登録されているか
		If ( isExists( envCurValue, envValue ) ) Then
			mb envValue & " is found"
			' envValueは登録されている
			' → なにもしない
		Else
			mb envValue & " is not found"
			' envValueが登録されていない
			' → 一番前に追加登録する
			envs.Item(envName) = envValue & ";" & envCurValue
		End If
	End If

End Sub

'----------------------------------------------------
' isExists
'
' allPathに、checkPathが含まれているかいないか調べる
'----------------------------------------------------
Function isExists(allPath, checkPath)
	mb allPath
	Dim returnValue
	returnValue = FALSE
	
	' ";"でallPathを区切って、すべて調べる
	Dim continue
	Dim startPos, foundPos
	Dim partPath
	continue = TRUE
	startPos = 1
	Do While ( continue )
		' ;を探す
		foundPos = InStr( startPos, allPath, ";", 1 )
		If ( ( foundPos = Null ) Or ( foundPos = 0 ) ) Then
			' みつからなかったら、終了
			continue = FALSE
		Else
			' みつかったら、startPos～(foundPos-1)がcheckPathでないかチェック
			partPath = Mid(allPath, startPos, (foundPos-startPos))
			mb partPath
			If ( StrComp( partPath, checkPath ) = 0 ) Then
				' checkPathがあった
				returnValue = TRUE
				' 検索終了
				continue = FALSE
			ELSE
				' checkPathがなかった
				' 次の文字から再検索
				startPos = foundPos + 1
			End If
			
		End If
	Loop
	
	' 全ループで、checkPathがなかった場合
	' 最後の;から末尾まで検索
	If ( Not returnValue ) Then
		' startPos～末尾を検索
		partPath = Mid(allPath, startPos, (Len(allPath)-startPos+1))
		mb partPath
		If ( StrComp( partPath, checkPath ) = 0 ) Then
			' checkPathがあった
			returnValue = TRUE
		End If
	End If
	
	isExists = returnValue
End Function


' デバッグ用MsgBox
' リリース時にFALSEにする
Sub mb(str)
	If ( FALSE ) Then
		MsgBox str
	End If
End Sub
