' UTF-8�̃t�@�C���𑀍삷��
' ����VBScrip�̃\�[�X�́Ashift-jis�Ƃ���
' UTF-8�ɂ���ɂ�wsf����Ăяo���K�v�����邽�߁A�߂�ǂ������B�i���A���΂ł���j

Option Explicit

Dim objFs
Set objFs = WScript.CreateObject("Scripting.FileSystemObject")

Dim testFileName
testFileName="utf-8test.txt"

' ------------------------------
' �V�K�쐬
Dim objTs
Dim OverWrite
OverWrite = True
' unicode�t�@�C�����쐬����ɂ́A��R������True��ݒ肷��
Set objTs = objFs.CreateTextFile(testFileName, OverWrite, True)
objTs.WriteLine "���{��̃e�X�g�P�Q�R"
objTs.WriteLine "���{��̃e�X�g�S�T�U"
objTs.Close
Set objTs = Nothing
' �� ���ꂾ�ƁAutf-16�ɂȂ�I<< NG >>

Dim objFile
Set objFile = objFs.GetFile(testFileName)
Dim ForWriting
ForWriting = 2
Set objTs = objFile.OpenAsTextStream(ForWriting, -1)
objTs.WriteLine "���{��̃e�X�g�P�Q�R"
objTs.WriteLine "���{��̃e�X�g�S�T�U"
objTs.Close
Set objTs = Nothing
' �� ����ł��Autf-16�ɂȂ�I<< NG! >>

' ���ꂪ��
Dim outStream
Set outStream = CreateObject("ADODB.Stream")
outStream.type = 2	' 1:�o�C�i���f�[�^ | 2:�e�L�X�g�f�[�^
msgbox outStream.mode
outStream.mode = 0	' 1:�ǂݎ�� | 2:�������� | 3:�ǂݎ��/�������ݗ���
' ���Ȃ���3����Ȃ��Ƃł��Ȃ��B
'   �܂��͎w�肵�Ȃ��Ă��ł���B�l��0�B
outStream.charset = "UTF-8"
outStream.open
outStream.WriteText "���{��̃e�X�g�P�Q�R", 0	' ��Q�����F0:��������������� | 1:������{���s��������������
outStream.WriteText "���{��̃e�X�g�S�T�U", 1
outStream.WriteText "���{��̃e�X�g�V�W�X", 1
' �ۑ�
outStream.SaveToFile testFileName, 2	' 1:�t�@�C�����Ȃ��ꍇ�̂ݍ쐬 | 2:����ꍇ�͏㏑��
outStream.close
Set outStream = Nothing

' ------------
' �ǉ���������
Dim addStream
Set addStream = CreateObject("ADODB.Stream")
addStream.type = 2
addStream.mode = 3	' 3:�ǂݎ��/�������ݗ���
'   �܂��͎w�肵�Ȃ��Ă��ł���B�l��0�B
addStream.charset = "UTF-8"
addStream.open
addStream.LoadFromFile testFileName
addStream.Position = addStream.Size		' �|�C���^���I�[��
addStream.WriteText "�ǉ��ł��P�Q�R", 1
addStream.SaveToFile testFileName, 2
addStream.close
Set addStream = Nothing



' ------------------------------
' �ǂݍ���

' �������悭�킩��Ȃ����ǁA����ς肱�̊֐�����UTF-8�͓ǂ߂Ȃ�
Set objTs = objFs.OpenTextFile(testFileName, 1,-1)
MsgBox objTs.ReadLine
MsgBox objTs.ReadLine
objTs.Close
Set objTs = Nothing

' ���ꂪ��
Dim inStream
Set inStream = CreateObject("ADODB.Stream")
inStream.type = 2
inStream.mode = 3
' ���Ȃ����R����Ȃ��Ƃł��Ȃ�
'   �܂��͎w�肵�Ȃ��Ă��ł���B�l��0�B
inStream.charset = "UTF-8"
inStream.open
inStream.LoadFromFile testFileName
Do While inStream.EOS = False
	MsgBox inStream.ReadText(-2)	' -1:�S���ǂݍ��� | -2:�P�s�ǂݍ���
Loop
inStream.close
Set inStream = Nothing

Set objFs = Nothing
