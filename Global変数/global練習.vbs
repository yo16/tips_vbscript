
Dim X         ' �O���[�o�� �X�R�[�v�� X ��錾���܂��B
X = "Global"      ' �O���[�o���� X �ɒl�������܂��B

Sub Proc1   ' �v���V�[�W����錾���܂��B
Dim X      ' ���[�J�� �X�R�[�v�� X ��錾���܂��B
X = "Local"   ' ���[�J���� X �ɒl�������܂��B
         ' �Ăяo������ X ���o�͂���v���V�[�W�����A
         ' Execute �X�e�[�g�����g�ō쐬���܂��B
         ' �O���[�o�� �X�R�[�v�Ɋ܂܂�邷�ׂĂ� Proc2 ��
         ' �p�����邽�߁A�O���[�o���� X ���o�͂���܂��B
  ExecuteGlobal "Sub Proc2: Print X: End Sub"
Print Eval("X")   ' ���[�J���� X ���o�͂��܂��B


Proc2      ' �O���[�o�� �X�R�[�v�� Proc2 ���Ăяo���ƁA
         ' "Global" ���������܂��B
End Sub

Proc2         ' Proc1 �̊O���� Proc2 ���g�p�ł��Ȃ����߁A
         ' ���̍s�ŃG���[���������܂��B
Proc1         ' Proc1 ���Ăяo���܂��B
  Execute "Sub Proc2: Print X: End Sub"
Proc2         ' Proc2 ���O���[�o���Ɏg�p�ł���悤��
         ' �Ȃ����̂ŁA���̌Ăяo���͐������܂��B


