Option Explicit

Dim objDict
Set objDict = CreateObject("Scripting.Dictionary")

objDict.CompareMode = vbTextCompare
'Add���\�b�h
' ��������Key�A��������Item
objDict.Add "1","����������"
objDict.Add "2","����������"
objDict.Add "3","����������"
objDict.Add "4","�����Ă�"

'Items���\�b�h��Keys���\�b�h���g���Ă݂�
Dim strItems,strKeys
strItems = objDict.Items
strKeys = objDict.Keys

Dim idx
'Count�v���p�e�B���g���Ă݂�
For idx = 0 To objDict.Count - 1
	MsgBox "�L�[ "&strKeys(idx)&" �ɑΉ�����f�[�^��"&strItems(idx)&"�ł��B"
Next

