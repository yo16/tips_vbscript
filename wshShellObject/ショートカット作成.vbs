' AllUserのデスクトップにショートカットを作成
' 2016/3/15 y.ikeda


Option Explicit

Dim WshShell
Set WshShell = WScript.CreateObject("WScript.Shell")
Dim strPutPath
strPutPath = WshShell.SpecialFolders("AllUsersDesktop")

Dim strLnkFilename, strLnkTagetPath, strWorkingDir, strDescription, strIconLocation, strPath
strLnkFilename   = "MEISTERSHiP-V2.7.lnk"
strLnkTagetPath  = "c:\CRESTAM\CSH\MEISTERSHiPV2.7.bat"
strWorkingDir    = "c:\CRESTAM\CSH\"
strDescription   = "MEISTERSHiP V2.7"
strIconLocation  = "c:\CRESTAM\LDM\ERBA.exe,0"
strPath = strPutPath + "\" + strLnkFilename

Dim oShellLink
Set oShellLink = WshShell.CreateShortcut(strPath)
oShellLink.TargetPath = strLnkTagetPath
oShellLink.WindowStyle = 1
oShellLink.Description = strDescription
oShellLink.IconLocation = strIconLocation
oShellLink.WorkingDirectory = strWorkingDir
oShellLink.Save

