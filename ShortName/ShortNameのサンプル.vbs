Option Explicit

msgbox ShowShortName("c:\winnt\notepad.exe")

Function ShowShortName(filespec)
   Dim fso, f, s
   Set fso = CreateObject("Scripting.FileSystemObject")
   Set f = fso.GetFile(filespec)
   s = UCase(f.Name) & "�� 8.3 �`���ł̃t�@�C�����́A"
   s = s & "���̂Ƃ���ł��B" & vbCrLf & f.ShortName 
   ShowShortName = s
End Function

