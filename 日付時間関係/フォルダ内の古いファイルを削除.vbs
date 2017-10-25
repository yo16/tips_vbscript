Option Explicit

Dim objFS
Set objFS = WScript.CreateObject("Scripting.FileSystemObject")

Dim objFolder
Set objFolder = objFS.GetFolder("削除テスト用フォルダ")

Dim colFiles
Set colFiles = objFolder.Files

Dim x,deleteCount
deleteCount = 0
For Each x in colFiles
	If deleteOldFile(x,5) Then deleteCount = deleteCount + 1
Next

MsgBox deleteCount & "個のファイルを削除しました！",,"★ 結果報告 ★"

'**************************************************
'関数[deleteOldFile]
'objFile	:削除する対象のファイルオブジェクト
'stockDay	:保存する期間(日にち単位)
'戻り値		:削除した場合はTrue
'			 削除しなかった場合はFalse
'
'[fileName]の最終更新日が[stockDay]日以上前の場合に
'削除する関数。
'**************************************************
Function deleteOldFile(objFile,stockDay)
	Dim dateDifference
	dateDifference = DateDiff("d",objFile.DateLastModified,Now)

	If (dateDifference >= stockDay) Then
		objFile.Delete
		deleteOldFile = True
	Else
		deleteOldFile = False
	End If
End Function


