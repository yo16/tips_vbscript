Option Explicit

msgbox ShowShortName("c:\winnt\notepad.exe")

Function ShowShortName(filespec)
   Dim fso, f, s
   Set fso = CreateObject("Scripting.FileSystemObject")
   Set f = fso.GetFile(filespec)
   s = UCase(f.Name) & "の 8.3 形式でのファイル名は、"
   s = s & "次のとおりです。" & vbCrLf & f.ShortName 
   ShowShortName = s
End Function

