Option Explicit


' [ �^ ]
' InputBox(prompt[, title][, default][, xpos][, ypos][, helpfile, context])
'
' prompt	: ���͂𑣂����b�Z�[�W
' title		: �E�B���h�E�̃^�C�g��
' default	: TextField�ɂ��炩���ߓ���镶����
' xpos		: ��ʂ̍��[����̋���(twip �P��)
' ypos		: ��ʂ̏�[����̋���(twip �P��)
' helpfile	: �w���v�t�@�C�������邱�Ƃ��ł���炵��
' context	: �w���v�t�@�C���̈����炵��


Dim modori

' ��{�I�ɂ��ꂭ�炢�킩���Ă�΂����񂶂��H
modori = InputBox("prompt", "title", "default")


' �J�N�j��
msgbox modori
