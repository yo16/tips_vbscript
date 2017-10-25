Option Explicit

Dim WMI, OS, Value, Shell

do while WScript.Arguments.Count = 0 and WScript.Version >= 5.7
    '##### WScript5.7 または Vista 以上かをチェック
    Set WMI = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\.\root\cimv2")
    Set OS = WMI.ExecQuery("SELECT *FROM Win32_OperatingSystem")
    For Each Value in OS
    if left(Value.Version, 3) < 6.0 then exit do
    Next

    '##### 管理者権限で実行
    Set Shell = CreateObject("Shell.Application")
    Shell.ShellExecute "wscript.exe", """" & WScript.ScriptFullName & """ uac", "", "runas"

    WScript.Quit
loop

'##### メイン処理を実行
WScript.Echo "Hello World"
