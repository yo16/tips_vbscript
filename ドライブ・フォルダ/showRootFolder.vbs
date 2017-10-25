msgbox showrootfolder("d")

Function ShowRootFolder(drvspec)
   Dim fso, f
   Set fso = CreateObject("Scripting.FileSystemObject")
   Set f = fso.GetDrive(drvspec)
   ShowRootFolder = f.RootFolder
End Function


