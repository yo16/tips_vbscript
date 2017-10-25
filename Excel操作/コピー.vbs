' �R�s�[��`�Ŏw�肵���Z���̃R�s�[�iFROM����TO�j
'
' �R�s�[��`�t�@�C���̃t�H�[�}�b�g�́A���L�B
' FROM�Z��,TO�Z����
' �擪��#�Ŏn�܂�s�̓R�����g
'
' 2007/10/12
Option Explicit

'** �ݒ� *******************
' �R�s�[���G�N�Z���t�@�C��
Dim strFromFileName
strFromFileName = "�R�s�[from.xls"

' �R�s�[��G�N�Z���t�@�C��
Dim strToFileName
strToFileName = "�R�s�[to.xls"

' �R�s�[��`�t�@�C��
Dim strDefFileName
strDefFileName = "�R�s�[��`.txt"


'** ���� *******************
'------------------------------
' ������
'------------------------------
' �t�@�C���V�X�e���I�u�W�F�N�g
Dim objFS,objFolder
Set objFS = CreateObject("Scripting.FileSystemObject")
Set objFolder = objFS.GetFolder(".")

'------------------------------
' �G�N�Z���t�@�C�����J��
'------------------------------
Dim objExcel
Dim objFromBook, objToBook
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = False
objExcel.DisplayAlerts = False
' From��Open
objExcel.Workbooks.Open objFolder.Path & "\" & strFromFileName
Set objFromBook = objExcel.ActiveWorkBook
' To��Open
objExcel.Workbooks.Open objFolder.Path & "\" & strToFileName
Set objToBook = objExcel.ActiveWorkBook


'------------------------------
' ��`�t�@�C����ǂ݂Ȃ���R�s�[
'------------------------------
Dim objTS
Set objTS = objFS.OpenTextFile( objFolder.Path & "\" & strDefFileName, 1 )
Dim strLine, aryLine, nLineNum
strLine = ""
nLineNum = 0
Dim nFromCol, nFromRow, nToCol, nToRow
Dim nRtn

Do Until objTS.AtEndOfStream
	nLineNum = nLineNum + 1
	strLine = objTS.ReadLine
	If ( Not( strLine = "" ) and ( Not(Left(strLine,1) = "#") ) ) Then
		
		aryLine = Split( strLine, "," )
		
		' From
		nRtn = GetXlsRowCol( aryLine(0), nFromRow, nFromCol )
		If ( nRtn < 0 ) Then
			MsgBox nLineNum & "�s�ڂŃG���[���������܂����B" & vbCrLf & strLine
			Quit -1
		End If
		
		' To
		nRtn = GetXlsRowCol( aryLine(1), nToRow, nToCol )
		If ( nRtn < 0 ) Then
			MsgBox nLineNum & "�s�ڂŃG���[���������܂����B" & vbCrLf & strLine
			Quit -1
		End If
		
		' �R�s�[
		objToBook.Sheets(1).Cells( nToRow, nToCol ).Value _
			= objFromBook.Sheets(1).Cells( nFromRow, nFromCol )
	End If
Loop
objTS.Close





'------------------------------
' �G�N�Z���t�@�C�������
'------------------------------
objFromBook.Close
objToBook.Save		' To�͕ۑ�
objToBook.Close

objExcel.Quit

Set objFromBook = Nothing
Set objToBook = Nothing

MsgBox "end��"




' �G�N�Z���̃Z�������s�ԍ��A��ԍ��ɕς���֐�
' �߂�l	0:����I��
' 2007/10/12
Function GetXlsRowCol( strXlsCellName, ByRef nXlsRow, ByRef nXlsCol )
	nXlsRow = 0
	nXlsCol = 0
	
	' ���K�\���I�u�W�F�N�g��ݒ�
	Dim regEx
	Set regEx = New RegExp
	regEx.Pattern = "^[A-Za-z]+[0-9]+$"
	
	' ����
	If ( Not regEx.Test( strXlsCellName ) ) Then
		' �A���}�b�`
		Msgbox "�t�H�[�}�b�g������Ă����ł�" & vbCrLf & strXlsCellName
		GetXlsRowCol = -1
		Exit Function
	End If
	
	' �啶����
	strXlsCellName = UCase( strXlsCellName )
	
	' �A���t�@�x�b�g�Ɛ����𕪂��邽�߂ɍČ���
	Dim nSepPos
	regEx.Pattern = "[0-9]+"
	Dim regMatches, regMatch
	Set regMatches = regEx.Execute( strXlsCellName )
	For Each regMatch in regMatches
		nSepPos = regMatch.FirstIndex
		Exit For
	Next
	
	' ������
	Dim strAlphabetPart, strNumberPart
	strAlphabetPart = Left( strXlsCellName, nSepPos )
	strNumberPart = Mid( strXlsCellName, nSepPos+1 )
	
	' �A���t�@�x�b�g�p�[�g�𐔒l�ɕϊ�����i26�i����10�i���j
	' ���̈ʂ���P�������ϊ����Ă���
	Dim nPos, i
	Dim cAZ, strAbc
	strAbc = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
	For i = 1 to Len( strAlphabetPart )
		' �P�����擾
		cAZ = Mid( strAlphabetPart, Len( strAlphabetPart )-i+1, 1 )
		' ���l��
		nPos = Instr( strAbc, cAZ )
		
		' �߂�l�i��j�֐ݒ�
		nXlsCol = nXlsCol + nPos * ( Len(strAbc) ^ (i-1) )
	Next
	' �߂�l�i�s�j�֐ݒ�
	nXlsRow = Int( strNumberPart )
	
	GetXlsRowCol = 0
End Function
