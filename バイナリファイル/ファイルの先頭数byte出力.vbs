Option Explicit
' �t�@�C���̐擪��byte���o��
' 2017 (c) yo16

Dim vbsTitle : vbsTitle = "head extract"

If WScript.Arguments.Count < 1 Then
	MsgBox "�Ώۃt�@�C����Drag & Drop���Ă�������", vbOkOnly , vbsTitle
	WScript.Quit 0
End If

' �h���b�v�������ׂẴt�@�C�����A�J��Ԃ�
Dim i
For i = 0 To WScript.Arguments.Count-1
	MakeHeadFileBin WScript.Arguments(i)
Next


' �w�肵���t�@�C���̐擪byte���o�͂���
Sub MakeHeadFileBin(inFilePath)
	'MsgBox inFilePath
	
	' �o�͂���byte��
	Dim outBinNum : outBinNum = 768	' 768(10)=300(16)
	
	
	Dim outFilePath : outFilePath = inFilePath & "_out.bin"
	
	Dim objBs : Set objBs = WScript.CreateObject("ADODB.Stream")
	objBs.Type = 1	' �o�C�i�����[�h
	objBs.Open
	objBs.LoadFromFile inFilePath
	Dim objOutBs : Set objOutBs = WScript.CreateObject("ADODB.Stream")
	objOutBs.Type = 1	' �o�C�i�����[�h
	objOutBs.Open
	
	objOutBs.Write objBs.Read(outBinNum)
	
	objOutBs.SaveToFile outFilePath, 2	' 2:�㏑��
	
	objOutBs.Close
	Set objOutBs = Nothing
	objBs.Close
	Set objBs = Nothing
	
	
	
End Sub
