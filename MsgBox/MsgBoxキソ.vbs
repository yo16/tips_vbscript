Option Explicit

'MsgBox�֐��̂���
'	2002/01/06

'�����������������킩��₷���B
'����ƁA���ɑ����Z����Ƃ����������B

' ������
' vbOKOnly				   0	[OK] �{�^���݂̂�\�����܂��B
' vbOKCancel			   1	[OK] �{�^���� [�L�����Z��] �{�^����\�����܂��B
' vbAbortRetryIgnore	   2	[���~]�A[�Ď��s]�A����� [����] �� 3 �̃{�^����\�����܂��B
' vbYesNoCancel			   3	[�͂�]�A[������]�A����� [�L�����Z��] �� 3 �̃{�^����\�����܂��B
' vbYesNo				   4	[�͂�] �{�^���� [������] �{�^����\�����܂��B
' vbRetryCancel			   5	[�Ď��s] �{�^���� [�L�����Z��] �{�^����\�����܂��B
' vbCritical			  16	�x�����b�Z�[�W �A�C�R����\�����܂��B
' vbQuestion			  32	�₢���킹���b�Z�[�W �A�C�R����\�����܂��B
' vbExclamation			  48	���Ӄ��b�Z�[�W �A�C�R����\�����܂��B
' vbInformation			  64	��񃁃b�Z�[�W �A�C�R����\�����܂��B
' vbDefaultButton1		   0	�� 1 �{�^����W���{�^���ɂ��܂��B
' vbDefaultButton2		 256	�� 2 �{�^����W���{�^���ɂ��܂��B
' vbDefaultButton3		 512	�� 3 �{�^����W���{�^���ɂ��܂��B
' vbDefaultButton4		 768	�� 4 �{�^����W���{�^���ɂ��܂��B
' vbApplicationModal	   0	�A�v���P�[�V���� ���[�_���ɐݒ肵�܂��B���b�Z�[�W �{�b�N�X�ɉ�������܂ŁA���ݑI�𒆂̃A�v���P�[�V�����̎��s���p���ł��܂���B
' vbSystemModal			4096	�V�X�e�� ���[�_���ɐݒ肵�܂��B���b�Z�[�W �{�b�N�X�ɉ�������܂ŁA���ׂẴA�v���P�[�V���������f����܂��B





'* �ϐ���` *
Dim nRtn		' Number�^�̕ϐ�

'* �֐��Ăяo�� *
nRtn = MsgBox("prompt!!", vbYesNoCancel + vbCritical, "title!!")

'* �߂�l�\�� *
MsgBox(nRtn)

' �߂�l
' �萔     �l �I�����ꂽ�{�^�� 
' vbOK     1  [OK] 
' vbCancel 2  [�L�����Z��] 
' vbAbort  3  [���~] 
' vbRetry  4  [�Ď��s] 
' vbIgnore 5  [����] 
' vbYes    6  [�͂�] 
' vbNo     7  [������] 

