Option Explicit

Dim objDic
Set objDic = CreateObject("Scripting.Dictionary")


If objDic.Exists("key1") Then
	objDic.Item("key1") = objDic.Item("key1") + 1
Else
	objDic.Add "key1", 1
End If

objDic.Add "key2", 1
If objDic.Exists("key2") Then
	objDic.Item("key2") = objDic.Item("key2") + 1
Else
	objDic.Add "key2", 1
End If


'msgbox objDic.Item("key1")
'msgbox objDic.Item("key2")

Dim key
For Each key In objDic.Keys
	msgbox key & ":" & objDic.Item(key)
Next
