Option Explicit

msgbox ShowShortName("sample.txt")
msgbox ShowShortPath("sample.txt")

Function ShowShortName(filespec)
   Dim fso, f, s
   Set fso = CreateObject("Scripting.FileSystemObject")
   Set f = fso.GetFile(filespec)
   s = UCase(f.Name) & "の 8.3 形式でのファイル名は、<BR>"   
   s = s & "次のとおりです。" & f.ShortName 
   ShowShortName = s
End Function


Function ShowShortPath(filespec)
   Dim fso, f, s
   Set fso = CreateObject("Scripting.FileSystemObject")
   Set f = fso.GetFile(filespec)
   s = UCase(f.Name) & "の 8.3 形式でのパス名は、<BR>"
   s = s & "次のとおりです。" & f.ShortPath 
   ShowShortPath = s
End Function
