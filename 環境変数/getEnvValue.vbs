Option Explicit

Dim iniFileName

Dim ParamName
'ParamName = "DBSID"
ParamName = InputBox ("�擾�������ϐ�������͂��Ă�������"," - * - * - ")

Dim DBName
if (getEnvValue((ParamName),DBName) = 0) then
	MsgBox(ParamName&" =>> "&DBName)
else
	MsgBox(ParamName&" �Ƃ����ϐ��͑��݂��܂���")
end if


'-------------'
' getEnvValue '
'-------------'

Function getEnvValue(byRef parameterName,parameterValue)
	Const ForReading = 1
	Const vbTextCompare = 1
	Dim iniFileName,rtnCode
	iniFileName = "environment.ini"

	'-- �I�u�W�F�N�g���� --'
	Dim iniFileObject
	Set iniFileObject = WScript.CreateObject("Scripting.FileSystemObject")
	Dim iniFileStream
	Set iniFileStream = iniFileObject.OpenTextFile(iniFileName,ForReading)

	'-- �t�@�C���ǂݍ��� --'
	Dim iniLine
	Dim eqArray
	Do While Not iniFileStream.AtEndOfStream
		iniLine = iniFileStream.ReadLine
		eqArray = Split(iniLine,"=",-1,vbTextCompare)
		If (UBound(eqArray) = 1) then
			If (Trim(eqArray(0)) = parameterName) then
				parameterValue = Trim(eqArray(1))
				getEnvValue = 0
				Exit Function
			End if
		End if
	Loop

	getEnvValue = -1
end Function




