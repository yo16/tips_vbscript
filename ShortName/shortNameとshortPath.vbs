Option Explicit

msgbox ShowShortName("sample.txt")
msgbox ShowShortPath("sample.txt")

Function ShowShortName(filespec)
   Dim fso, f, s
   Set fso = CreateObject("Scripting.FileSystemObject")
   Set f = fso.GetFile(filespec)
   s = UCase(f.Name) & "�� 8.3 �`���ł̃t�@�C�����́A<BR>"   
   s = s & "���̂Ƃ���ł��B" & f.ShortName 
   ShowShortName = s
End Function


Function ShowShortPath(filespec)
   Dim fso, f, s
   Set fso = CreateObject("Scripting.FileSystemObject")
   Set f = fso.GetFile(filespec)
   s = UCase(f.Name) & "�� 8.3 �`���ł̃p�X���́A<BR>"
   s = s & "���̂Ƃ���ł��B" & f.ShortPath 
   ShowShortPath = s
End Function
