Option Explicit
' �ړ�
' 2010/06/23


' MoveFile���g��



Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")
fso.MoveFile "rename_1.txt", "rename_2.txt"

' Move��̃t�@�C��������ꍇ�́A�G���[�ɂȂ�

' �t�H���_���Ȃ��ꍇ�͂ǂ��Ȃ�񂾂낤���B
' �� ���s���G���[���o��

