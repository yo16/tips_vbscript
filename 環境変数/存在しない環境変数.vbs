Option Explicit

	Dim WshShell
	Set WshShell = WScript.CreateObject ("WScript.Shell")


	Dim WshEnv
	Set WshEnv = WshShell.Environment("VOLATILE")


	'���݂��Ȃ����ϐ�������͂���ƁE�E�E
	If (WshEnv.Item("ABCDEFG") = "") Then
		msgbox "item is (����0�̕�����)"
	else
		if (WshEnv.Item("ABCDEFG") = Null) Then
			msgbox "item is Null"
		else
			msgbox "item is ?"
		end if
	end if
	


