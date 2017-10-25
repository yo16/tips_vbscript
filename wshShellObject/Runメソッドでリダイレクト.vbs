Option Explicit

Dim WshShell
Set WshShell = WScript.CreateObject("WScript.Shell")

'‘æˆêˆø”:exeƒtƒ@ƒCƒ‹–¼
'‘æ“ñˆø”:‚P¨‰ÂŽ‹ ‚O¨•s‰ÂŽ‹
'‘æŽOˆø”:‚P¨I—¹‚ð‘Ò‚Â ‚O¨‘Ò‚½‚¸‚ÉŽŸ‚ðŽÀs
'–ß‚è’l  :‚O¨³íI—¹ ‚P¨ˆÙíI—¹
Dim runRtn
runRtn = WshShell.Run("cmd /C echo %date% %time% > date.txt",1,1)

' del‚Æ“¯‚¶‚â‚è•û

msgbox runRtn
