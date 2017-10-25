'----------------------------------------------------
' GetRegBrock
' DiplayNameを検索して、そのブロックの最初と最後の行番号を返す
'
' return: 0:正常終了
'        -1:
' 
' 注意
' ・ブロックに"UninstallString"が含まれていることを以って
'   生きたブロックとみなす
'----------------------------------------------------
Function GetRegBrock( ByRef pDisplayName, pFileName, pBeginLine, pEndLine )
	
	Dim objFS, objTS
	Set objFS = WScript.CreateObject("Scripting.FileSystemObject")
	
	' ファイルオープン
	
	' 
	' 空行を記憶
	
	
End Function
