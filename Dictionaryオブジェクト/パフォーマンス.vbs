Option Explicit

' �n�b�V���̃p�t�H�[�}���X����
' 2016/10/20 y.ikeda
' �o�^�������ɑ΂��āA�������x���ǂ̒��x�x���Ȃ邩


' �e�X�g������
Dim testSearchNumber
testSearchNumber = 1000000


' �o�^����ς��āA���x���v��
TestUseDictionary 100000
' �� 7s
TestUseDictionary 200000
' �� 12s
TestUseDictionary 300000
' �� 17s
TestUseDictionary 400000
' �� 22s

' ����ȂɈ����Ȃ����
' �ɒ[�Ɉ������邱�Ƃ��Ȃ�





Sub TestUseDictionary( registNumber )
	Dim objDic
	Dim i
	Dim startDt, endDt
	Dim spentTime_s
	Dim testKeyNo
	Dim testItem
	Set objDic = CreateObject("Scripting.Dictionary")
	For i=0 To registNumber
		objDic.Add "KEY-"&i, "ITEM-"&i
	Next

	Randomize

	' �v���J�n
	startDt = Now

	For i=0 To testSearchNumber
		' 0�`registNumber-1�̃����_���Ȓl��ݒ�
		testKeyNo = Int( registNumber*Rnd )
		testItem = objDic.Item( "KEY-"&testKeyNo )
	Next

	' �v���I��
	endDt = Now
	spentTime_s = DateDiff("s", startDt, endDt)

	MsgBox "�o�^��:" & registNumber & vbCrLf & _
		spentTime_s & "(s)"

End Sub
