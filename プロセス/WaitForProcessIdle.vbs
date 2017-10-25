Option Explicit
' http://yozda.exblog.jp/15444817/

If WaitForProcessIdle("CADmeister.exe") Then MsgBox "アイドル状態" Else MsgBox "ない"



' WaitForProcessIdle
' プロセスが終了するまで待つ
' 戻り値
' True  : 指定プロセスがアイドル状態になった(or 終了した)
' False : 指定プロセスが存在しない
Function WaitForProcessIdle(ProcessName)
	'http://www.tek-tips.com/viewthread.cfm?qid=395765
	
	' 戻り値を初期設定
	WaitForProcessIdle = False
	
	' プロセスIDを取得する
	Dim Process
	Dim pID: pID = 0
	For Each Process in GetObject("winmgmts:").ExecQuery("Select * from Win32_Process where Name = '" & ProcessName & "'")
		pID = Process.Handle
		Exit For
	Next
	If pID = 0 Then Exit Function ' プロセスが見つからない
	On Error Resume Next
	WScript.StdOut.Write ProcessName &"(" & pID &")"
	
	
	Dim cmd
	cmd = "Select * from Win32_PerfRawData_PerfProc_Process where IDProcess = '" & pID & "'"
	Dim objService
	Set objService = GetObject("Winmgmts:{impersonationlevel=impersonate}!\Root\Cimv2")   
	Dim objInstance, n1, d1
	For Each objInstance in objService.ExecQuery(cmd)
		n1 = objInstance.PercentProcessorTime
		d1 = objInstance.TimeStamp_Sys100NS
		Exit For
	Next
	Dim n0, d0, cpuusage
	Do
		If objInstance.Name = "" Then Exit Do ' プロセスが終了
		n0 = n1
		d0 = d1
		WScript.Sleep(1000)
		WScript.StdOut.Write "."
		For Each objInstance in objService.ExecQuery(cmd)
			n1 = objInstance.PercentProcessorTime
			d1 = objInstance.TimeStamp_Sys100NS
			Exit For
		Next
		cpuusage = Round((n1 - n0)/(d1 - d0)*100, 0)
	Loop Until cpuusage = 0   ' CPU使用率が0になるまでループ
	WScript.StdOut.Write vbCrLf
	
	On Error GoTo 0
	WaitForProcessIdle = True
End Function
