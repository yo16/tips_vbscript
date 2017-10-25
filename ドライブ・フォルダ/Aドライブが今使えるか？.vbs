Option Explicit


Dim fso, d, s, t


Set fso = CreateObject("Scripting.FileSystemObject")
Set d = fso.GetDrive("A")
Select Case d.DriveType
   Case 0:t = "不明"
   Case 1:t = "リムーバブル ディスク"
   Case 2:t = "ハード ディスク"
   Case 3:t = "ネットワーク ドライブ"
   Case 4:t = "CD-ROM"
   Case 5:t = "RAM ディスク"
End Select
s = "ドライブ " & d.DriveLetter & ": - " & t
If d.IsReady Then 
   s = s & vbCrLf & "ドライブの準備ができています。"
Else
   s = s & vbCrLf & "ドライブの準備ができていません。"
End If

MsgBox s
