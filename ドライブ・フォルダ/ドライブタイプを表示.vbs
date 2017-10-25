's045.vbs

Option Explicit

Dim modori
modori = inputbox("driveName")

Dim objFS, objDrive
' FileSystemObject オブジェクトを生成する
Set objFS = WScript.CreateObject("Scripting.FileSystemObject")
' Drive オブジェクトを生成する
Set objDrive = objFS.GetDrive(modori)
' ドライブタイプを表示する
WScript.Echo objDrive.DriveType
