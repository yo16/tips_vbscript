' ���t�^�̃L�\

Dim v_dateA, v_dateB
Dim v_diffAB



' ���������t�^�ɕϊ�
v_dateA = CDate("2004/4/1")
v_dateB = CDate("2004/5/1")

msgbox v_dateA
msgbox v_dateB


' ���낢��Ȃ����ŁA�Q�̍�������Ă݂�
v_diffAB = DateDiff("y", v_dateA, v_dateB)		' �N�ԒʎZ�� �`�a
msgbox v_diffAB
v_diffAB = DateDiff("d", v_dateA, v_dateB)		' ���t �`�a
msgbox v_diffAB
v_diffAB = DateDiff("y", v_dateB, v_dateA)		' �N�ԒʎZ�� �a�`
msgbox v_diffAB
v_diffAB = DateDiff("d", v_dateB, v_dateA)		' ���t �a�`
msgbox v_diffAB

' ���ʁ� ���Ԃ͊֌W����B��R�����|��Q������Ԃ��I


