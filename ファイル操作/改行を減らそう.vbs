Option Explicit



Dim YNmodori
YNmodori = MsgBox("改行を減らしてもいいですか？",4,"改行を減らそう！")
If (YNmodori <> 6) Then
	WScript.Quit
End If




Dim objFS,objFolder,colFiles
Dim x
' FileSystemObject オブジェクトを生成する
Set objFS = WScript.CreateObject("Scripting.FileSystemObject")
' カレントフォルダのFolderオブジェクトを取得する
Set objFolder = objFS.GetFolder(".")
' カレントフォルダに含まれるすべてのファイルを取得する
Set colFiles = objFolder.Files






Dim objTS,workTS,workFile
Dim lineStr,tmpLineStr
Dim nullCount,xName
Dim fileCount
fileCount = -1
For Each x in colFiles
	fileCount = fileCount + 1
	If Not (x.Name = WScript.ScriptName) Then
		Set objTS = x.OpenAsTextStream
		Set workTS = objFS.CreateTextFile("改行を減らそうwork.txt",TRUE)
		nullCount = 0
		Do Until objTS.AtEndOfStream
			lineStr = objTS.ReadLine
			tmpLineStr = Replace(lineStr,vbTab,"")
			tmpLineStr = Trim(tmpLineStr)
			If (tmpLineStr = "") Then
				nullCount = nullCount + 1
				If Not (nullCount >= 3) Then
					workTS.WriteBlankLines(1)
				End If
			Else
				nullCount = 0
				workTS.WriteLine(RTrim(lineStr))
			End If
		Loop
		objTS.Close
		workTS.Close

		xName = x.Name
		objFS.DeleteFile xName
		Set workFile = objFS.GetFile("改行を減らそうwork.txt")
		workFile.Name = xName
	End If
Next




msgbox "無事終了～★",,fileCount&"個のファイルを見ましたよ"





