Option Explicit

msgbox ToggleArchiveBit("sample.txt")

Function ToggleArchiveBit(filespec)
   Dim fso, f
   Set fso = CreateObject("Scripting.FileSystemObject")
   Set f = fso.GetFile(filespec)
   If f.attributes and 32 Then
      f.attributes = f.attributes - 32
      ToggleArchiveBit = "Archive �r�b�g���I�t�ɂ��܂����B"
   Else
      f.attributes = f.attributes + 32
      ToggleArchiveBit = "Archive �r�b�g���I���ɂ��܂����B"
   End If
End Function

