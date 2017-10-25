Option Explicit

Dim strCellName
Dim nRow, nCol
Dim nRtn

strCellName = "aa3"
nRtn = GetXlsRowCol(strCellName, nRow, nCol)
MsgBox nRow & "-" & nCol


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
