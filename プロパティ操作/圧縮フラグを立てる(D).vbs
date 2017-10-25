Option Explicit



Dim YNmodori
YNmodori = MsgBox("圧縮フラグを立ててもいいですか？",4,"圧縮フラグを立てる")
If (YNmodori <> 6) Then
	WScript.Quit
End If



msgbox compressFolder("a")


'********************************************
'関数:compressFolder
'引数:folderName:圧縮するファイル名
'
'＊＊説明＊＊
'ファイルのプロパティ[圧縮(M)]をチェックする
'********************************************
Function compressFolder(folderName)
	'ファイルのプロパティを取得する
	Dim objFS,objFolder
	Set objFS = WScript.CreateObject("Scripting.FileSystemObject")
	Set objFolder = objFS.GetFolder(folderName)
	Dim propertyValue
	propertyValue = objFolder.Attributes

msgbox "property"&propertyValue
	'明らかに圧縮されていない場合は圧縮フラグを立てて正常終了
	If (propertyValue < 2048) Then
		objFolder.Attributes = propertyValue + 2048
		compressFolder = 0
		Exit Function
	End If

	'プロパティの値を２進数にする
	Dim sho,amari,idx,propertyValue_2
	sho = propertyValue
	idx = 0
	Do Until (sho = 1)
		nishinNumber = nishinNumber + ( (sho mod 2) * (10^idx) )
		sho = sho \ 2
		idx = idx + 1
	Loop
	propertyValue_2 = propertyValue_2 + 10^idx

	'圧縮の状態を示すフラグを取得する
	Dim compressFlg
	compressFlg = Left(Right(propertyValue_2,11),1)

	'圧縮されていない場合は圧縮フラグを立てる
	If (compressFlg = "0") Then
		objFolder.Attributes = propertyValue + 2048
	End If

	compressFolder = 0

End Function


