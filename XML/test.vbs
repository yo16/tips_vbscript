Option Explicit

Dim objDOM, rtResult

Set objDOM = WScript.CreateObject("MSXML2.DOMDocument")
rtResult = objDOM.load("test.xml")
If rtResult = True Then
	procDispDatas objDOM.childNodes
End If
Set objDOM = Nothing


Sub procDispDatas(objNode)
	Dim obj
	For Each obj In objNode
		If obj.nodeType = 3 and obj.parentNode.nodeName = "title" Then
			MsgBox obj.parentNode.nodeName & " : " & obj.nodeValue
		End If
		If obj.hasChildNodes Then
			procDispDatas obj.childNodes
		End If
	Next
End Sub 

