Option Explicit

' �u��
' 2006/10/11 ikeda


msgbox Replace("XXpXXPXXp", "p", "Y")   ' ������̍ŏ�����A�o�C�i�� ���[�h�Ŕ�r���s���܂��B"XXYXXPXXY"��Ԃ��܂��B
msgbox Replace("XXpXXPXXp", "p", "Y", 3, -1, 1)   ' 3 �Ԗڂ̈ʒu����e�L�X�g ���[�h�Ŕ�r���s���܂��B"YXXYXXY" ��Ԃ��܂��B

' �q�b�g���Ȃ��Ƃ��͉������Ȃ�
MsgBox Replace("abcde", "x", "y")
