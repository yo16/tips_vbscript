Option Explicit

msgbox ToggleArchiveBit("sample.txt")

Function ToggleArchiveBit(filespec)
   Dim fso, f
   Set fso = CreateObject("Scripting.FileSystemObject")
   Set f = fso.GetFile(filespec)
   If f.attributes and 1 Then
      f.attributes = f.attributes - 1
      ToggleArchiveBit = "ReadOnly ビットをオフにしました。"
   Else
      f.attributes = f.attributes + 1
      ToggleArchiveBit = "ReadOnly ビットをオンにしました。"
   End If
End Function

