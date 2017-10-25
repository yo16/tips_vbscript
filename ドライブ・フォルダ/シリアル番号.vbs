Option Explicit


Dim objFS,objDrive
Dim strSerialNumber

Set objFS = WScript.CreateObject("Scripting.FileSystemObject")
Set objDrive = objFS.GetDrive("C")

strSerialNumber = Hex(objDrive.SerialNumber)

MsgBox Left(strSerialNumber,4) & "-" & Right(strSerialNumber,4),,"SerialNumber(C)"

Set objDrive = objFS.GetDrive("E")

strSerialNumber = Hex(objDrive.SerialNumber)

MsgBox Left(strSerialNumber,4) & "-" & Right(strSerialNumber,4),,"SerialNumber(E)"
