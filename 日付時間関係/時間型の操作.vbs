Option Explicit

Dim nowNichiji
nowNichiji = Now

'MsgBox nowNichiji+1'<=1���v���X(�킩��Â炢)

nowNichiji = DateAdd("m",1,nowNichiji)
MsgBox nowNichiji
