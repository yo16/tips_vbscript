' ��A123
' ���N���|123
' �g�����|123��S�������ł���悤�ȃ}�b�`������

Dim str1 : str1 = "��A123"
Dim str2 : str2 = "���N���|456"
Dim str3 : str3 = "�g�����|789"



'���K�\���I�u�W�F�N�g���쐬
Dim regEx
Set regEx = New RegExp

'�p�^�[����ݒ�
regEx.Pattern = "((��A)|(���N���|)|(�g�����|))([0-9]+)"

'Matches�R���N�V�����ɓ���
Dim Matches
Set Matches = regEx.Execute(str1)
MsgBox "str1.Count:" & Matches.Item(0).Value

Set Matches = regEx.Execute(str2)
MsgBox "str2.Count:" & Matches.Item(0).Value

Set Matches = regEx.Execute(str3)
MsgBox "str3.Count:" & Matches.Item(0).Value
MsgBox Matches.Item(0).SubMatches(4)
