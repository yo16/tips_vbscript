Option Explicit


Dim fso, d, s, t


Set fso = CreateObject("Scripting.FileSystemObject")
Set d = fso.GetDrive("A")
Select Case d.DriveType
   Case 0:t = "�s��"
   Case 1:t = "�����[�o�u�� �f�B�X�N"
   Case 2:t = "�n�[�h �f�B�X�N"
   Case 3:t = "�l�b�g���[�N �h���C�u"
   Case 4:t = "CD-ROM"
   Case 5:t = "RAM �f�B�X�N"
End Select
s = "�h���C�u " & d.DriveLetter & ": - " & t
If d.IsReady Then 
   s = s & vbCrLf & "�h���C�u�̏������ł��Ă��܂��B"
Else
   s = s & vbCrLf & "�h���C�u�̏������ł��Ă��܂���B"
End If

MsgBox s
