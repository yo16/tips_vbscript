Option Explicit

msgbox ToggleArchiveBit("sample.txt")

Function ToggleArchiveBit(filespec)
   Dim fso, f
   Set fso = CreateObject("Scripting.FileSystemObject")
   Set f = fso.GetFile(filespec)
   If f.attributes and 1 Then
      f.attributes = f.attributes - 1
      ToggleArchiveBit = "ReadOnly �r�b�g���I�t�ɂ��܂����B"
   Else
      f.attributes = f.attributes + 1
      ToggleArchiveBit = "ReadOnly �r�b�g���I���ɂ��܂����B"
   End If
End Function

