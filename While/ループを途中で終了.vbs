Option Explicit

Dim idx
idx = 1
Dim str
str = ""
Do While (idx < 10)
	str = str & idx
	
	' �r���ŏI��
	If ( idx = 5 ) Then
		Exit Do
		' Exit �́A
		' Do...Loop ���[�v�AFor...Next ���[�v�AFunction �v���V�[�W���܂��� Sub �v���V�[�W�����甲���o�����߂̃t���[����X�e�[�g�����g�ł��B

		
	End If
	
	
	idx = idx + 1
Loop

msgbox str

