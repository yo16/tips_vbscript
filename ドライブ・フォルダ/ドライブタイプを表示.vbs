's045.vbs

Option Explicit

Dim modori
modori = inputbox("driveName")

Dim objFS, objDrive
' FileSystemObject �I�u�W�F�N�g�𐶐�����
Set objFS = WScript.CreateObject("Scripting.FileSystemObject")
' Drive �I�u�W�F�N�g�𐶐�����
Set objDrive = objFS.GetDrive(modori)
' �h���C�u�^�C�v��\������
WScript.Echo objDrive.DriveType
