Set WshShell = WScript.CreateObject("WScript.Shell")

Set WshSysEnv = WshShell.Environment("PROCESS")

Dim p_i
p_i = 0
for each strenv in WshSysEnv
p_i = p_i + 1
next
msgbox "count -> "&p_i

