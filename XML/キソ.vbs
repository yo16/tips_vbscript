Option Explicit

Dim objDOM, rtResult

Set objDOM = WScript.CreateObject("MSXML2.DOMDocument")
'rtResult = objDOM.load("test.xml")
rtResult = objDOM.load("memsone8.0.xml")
If rtResult = True Then
	EchoNodeName 1, 2, objDOM.childNodes
	
End If
Set objDOM = Nothing

msgbox "end"



Sub EchoNodeName(curFloor, lastFloor, objNode)
	
	Dim obj
	For Each obj In objNode
		
		MsgBox curFloor & ":" & obj.nodeName
		
		' �q����������q�����\���ilastFloor�܂Łj
		If ( curFloor < lastFloor ) Then
			If obj.hasChildNodes Then
				EchoNodeName curFloor+1, lastFloor, obj.childNodes
			End If
		End If
	Next
	
End Sub


