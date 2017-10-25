Option Explicit

Dim objWshNetwork
Set objWshNetwork = WScript.CreateObject("WScript.Network")

MsgBox objWshNetwork.ComputerName,,"ComputerName"

MsgBox objWshNetwork.UserDomain,,"UserDoamain"

MsgBox objWshNetwork.UserName,,"UserName"


