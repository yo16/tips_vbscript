' �L�\

Dim str : str = "abc123xxx456def789"



'���K�\���I�u�W�F�N�g���쐬
Dim regEx
Set regEx = New RegExp

'�p�^�[����ݒ�
regEx.Pattern = "([0-9])[a-z]"
'������S�̂���������悤�ɐݒ�
regEx.Global = True

'Matches�R���N�V�����ɓ���
Dim Matches
Set Matches = regEx.Execute(str)

'1���\��
Dim Match, subMatch
For Each Match in Matches
	' �}�b�`�����S�̂́AMatch.Value�ɓ����Ă���
	MsgBox Match.Value
	
	' ()���̕��������肽���ꍇ�́ASubMatches���g��
	For Each subMatch in Match.SubMatches
		MsgBox subMatch
		' .Value�łȂ��̂Œ���
	Next
	
Next

' �}�b�`��
MsgBox "�}�b�`���F" & Matches.Count

' �ŏ��̂P�����g�p���Ȃ���΁AItem���g����
MsgBox Matches.Item(0).Value
