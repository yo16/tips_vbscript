Option Explicit

' 1,2,3,4,err!
' Resume Nextがあると、そのままスルーするが
' if errには入る。


msgbox "1"
'Err.Raise(1)
test1
msgbox "4"

if err then msgbox "err!" & err.number


Sub test1
	On Error Resume Next
	msgbox "2"
	Err.Raise(2)
	msgbox "3"
End Sub

