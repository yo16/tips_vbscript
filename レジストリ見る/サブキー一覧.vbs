Option Explicit
' インポート
Execute ReadFile("RegClass.vbs")
Execute ReadFile("RegClassCtl.vbs")


Dim WshShell
Set WshShell = WScript.CreateObject("WScript.Shell")

' レジストリエクスポート
Dim regKeyStr, regExpFile
regKeyStr = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
regExpFile = "export.txt"

'第一引数:exeファイル名
'第二引数:１→可視 ０→不可視
'第三引数:１→終了を待つ ０→待たずに次を実行
'戻り値  :０→正常終了 １→異常終了
Dim rtn
rtn = WshShell.Run( "reg export " & regKeyStr & " " & regExpFile, 0, 1 )


' レジストリを読んで、情報を整理
Dim regObj
Set regObj = GetRegClass(regExpFile, regKeyStr)

' レジストリオブジェクトから、情報を取得
Dim installPath_MemsONE, installPath_Mz
Dim subKeyCount
subKeyCount = regObj.GetSubKeysCount()
Dim valueStr
Dim i
For i = 0 to subKeyCount-1
	valueStr = regObj.GetSubKeyObjAt(i).GetValueByName( """DisplayName""" )
	If ( valueStr = """MemsONE""" ) Then
		installPath_MemsONE = regObj.GetSubKeyObjAt(i).GetValueByName( """InstallLocation""" )
		installPath_MemsONE = Replace(installPath_MemsONE, """", "")
		installPath_MemsONE = Replace(installPath_MemsONE, "\\", "\")
	ElseIf ( valueStr = """SMART PP 1.4""" ) Then
		installPath_Mz = regObj.GetSubKeyObjAt(i).GetValueByName( """InstallLocation""" )
		installPath_Mz = Replace(installPath_Mz, """", "")
		installPath_Mz = Replace(installPath_Mz, "\\", "\")
	End If
	
Next

msgbox installPath_MemsONE
msgbox installPath_Mz






' 外部ファイルを読み込む（インポート用）
Function ReadFile(ByVal FileName)
	Const ForReading = 1
	
	Dim FileShell
	Set FileShell = WScript.CreateObject("Scripting.FileSystemObject")
	
	ReadFile = FileShell.OpenTextFile(FileName, ForReading, False).ReadAll()
End Function
