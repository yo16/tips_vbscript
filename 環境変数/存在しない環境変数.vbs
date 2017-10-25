Option Explicit

	Dim WshShell
	Set WshShell = WScript.CreateObject ("WScript.Shell")


	Dim WshEnv
	Set WshEnv = WshShell.Environment("VOLATILE")


	'‘¶İ‚µ‚È‚¢ŠÂ‹«•Ï”–¼‚ğ“ü—Í‚·‚é‚ÆEEE
	If (WshEnv.Item("ABCDEFG") = "") Then
		msgbox "item is (’·‚³0‚Ì•¶š—ñ)"
	else
		if (WshEnv.Item("ABCDEFG") = Null) Then
			msgbox "item is Null"
		else
			msgbox "item is ?"
		end if
	end if
	


