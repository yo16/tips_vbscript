Option Explicit



Dim YNmodori
YNmodori = MsgBox("���s�����炵�Ă������ł����H",4,"���s�����炻���I")
If (YNmodori <> 6) Then
	WScript.Quit
End If




Dim objFS,objFolder,colFiles
Dim x
' FileSystemObject �I�u�W�F�N�g�𐶐�����
Set objFS = WScript.CreateObject("Scripting.FileSystemObject")
' �J�����g�t�H���_��Folder�I�u�W�F�N�g���擾����
Set objFolder = objFS.GetFolder(".")
' �J�����g�t�H���_�Ɋ܂܂�邷�ׂẴt�@�C�����擾����
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
		Set workTS = objFS.CreateTextFile("���s�����炻��work.txt",TRUE)
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
		Set workFile = objFS.GetFile("���s�����炻��work.txt")
		workFile.Name = xName
	End If
Next




msgbox "�����I���`��",,fileCount&"�̃t�@�C�������܂�����"





