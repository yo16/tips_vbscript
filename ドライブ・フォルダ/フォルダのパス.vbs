Option Explicit

Dim objFs
Set objFs = WScript.CreateObject("Scripting.FileSystemObject")

Dim objDir
Set objDir = objFs.GetFolder(".")

Msgbox objDir.Path
Msgbox objDir.Name
