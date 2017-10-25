Option Explicit

msgbox ToggleArchiveBit("sample.txt")

Function ToggleArchiveBit(filespec)
   Dim fso, f
   Set fso = CreateObject("Scripting.FileSystemObject")
   Set f = fso.GetFile(filespec)
   If f.attributes and 32 Then
      f.attributes = f.attributes - 32
      ToggleArchiveBit = "Archive ビットをオフにしました。"
   Else
      f.attributes = f.attributes + 32
      ToggleArchiveBit = "Archive ビットをオンにしました。"
   End If
End Function

