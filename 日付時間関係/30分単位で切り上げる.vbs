' 30���P�ʂŐ؂�グ��֐�
' 2007/10/31
Option Explicit




MsgBox GetKiriage30("9:15")
MsgBox GetKiriage30("9:45")
MsgBox GetKiriage30("0:00")
MsgBox GetKiriage30("9:30")
MsgBox GetKiriage30("10:00")
MsgBox GetKiriage30("24:00")
MsgBox GetKiriage30("23:30")
MsgBox GetKiriage30("24:59")





' �R�O���P�ʂŐ؂�グ��
' ������̃t�H�[�}�b�g�́A��:��
Function GetKiriage30( strBaseTime )
	Dim aryMS, strM, strS
	aryMS = Split( strBaseTime, ":" )
	If (UBound(aryMS) < 1) Then
		GetKiriage30 = ""
		Exit Function
	End If
	
	strM = aryMS(0)
	strS = aryMS(1)
	
	Dim nM, nS
	nM = Int(strM)
	nS = Int(strS)
	
	If ( (nS mod 30) = 0 ) Then
		' ���܂聁�O���؂�グ��K�v�Ȃ����Ȃɂ����Ȃ�
		
	Else
		If ( nS > 30 ) Then
			nM = nM + 1
			
			strM = CStr( nM )
			strS = "00"
		Else
			strS = "30"
		End If
		
	End If
	
	
	GetKiriage30 = strM & ":" & strS
End Function

