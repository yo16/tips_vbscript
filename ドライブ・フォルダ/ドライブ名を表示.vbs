's056.vbs

Option Explicit
Dim objFS, objDrives
Dim strDrives, x
' FileSystemObject �I�u�W�F�N�g�𐶐�����
Set objFS = WScript.CreateObject("Scripting.FileSystemObject")
Set objDrives = objFS.Drives
For Each x in objDrives
	strDrives = strDrives & x & vbCRLF
Next
WScript.Echo strDrives
