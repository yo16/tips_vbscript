Option Explicit
' ファイルの先頭数byteを出力
' 2017 (c) yo16

Dim vbsTitle : vbsTitle = "head extract"

If WScript.Arguments.Count < 1 Then
	MsgBox "対象ファイルをDrag & Dropしてください", vbOkOnly , vbsTitle
	WScript.Quit 0
End If

' ドロップしたすべてのファイル分、繰り返す
Dim i
For i = 0 To WScript.Arguments.Count-1
	MakeHeadFileBin WScript.Arguments(i)
Next


' 指定したファイルの先頭byteを出力する
Sub MakeHeadFileBin(inFilePath)
	'MsgBox inFilePath
	
	' 出力するbyte数
	Dim outBinNum : outBinNum = 768	' 768(10)=300(16)
	
	
	Dim outFilePath : outFilePath = inFilePath & "_out.bin"
	
	Dim objBs : Set objBs = WScript.CreateObject("ADODB.Stream")
	objBs.Type = 1	' バイナリモード
	objBs.Open
	objBs.LoadFromFile inFilePath
	Dim objOutBs : Set objOutBs = WScript.CreateObject("ADODB.Stream")
	objOutBs.Type = 1	' バイナリモード
	objOutBs.Open
	
	objOutBs.Write objBs.Read(outBinNum)
	
	objOutBs.SaveToFile outFilePath, 2	' 2:上書き
	
	objOutBs.Close
	Set objOutBs = Nothing
	objBs.Close
	Set objBs = Nothing
	
	
	
End Sub
