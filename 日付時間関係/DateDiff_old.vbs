Option Explicit
'-----------------------------------------------------
' DataDiff�̃e�X�g
'-----------------------------------------------------

' ����
Dim v_Today
v_Today = Date

' �w���
Dim v_Aruhi
v_Aruhi = CDate("2006/5/30")


Dim v_diffDays
v_diffDays = DateDiff( "y", v_Today, v_Aruhi )
' �P�ڂ̈���
' �ݒ�l	���e 
' yyyy		�N 
' q			�l���� 
' m			�� 
' y			�N�ԒʎZ�� 
' d			�� 
' w			�T�� 
' ww		�T 
' h			�� 
' n			�� 
' s			�b 





msgbox( v_diffDays )

