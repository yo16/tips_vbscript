Option Explicit

Dim objFS,objFolder
Set objFS = CreateObject("Scripting.FileSystemObject")
Set objFolder = objFS.GetFolder(".")
Msgbox "D&D�̂Ƃ��́A�J�����g�t�H���_���ς��܂�" & vbCrLf & objFolder.Path



' �����Ă��̑΍�
dim objShell
set objShell = CreateObject("WScript.Shell")
objShell.CurrentDirectory = objFS.GetParentFolderName(WScript.ScriptFullName)


Set objFolder = objFS.GetFolder(".")
Msgbox "�Q�x�ڂ́H" & vbCrLf & objFolder.Path
