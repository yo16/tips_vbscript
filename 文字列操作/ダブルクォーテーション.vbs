Option Explicit


'�Q�A���ŏ����ƂP��"�ƂȂ�
msgbox """"

'�ϐ��ł̑���
Dim testStr
testStr = """abc"""
msgbox testStr

'�t�@�C���̑���
Dim objFS,objTS
Set objFS = WScript.CreateObject("Scripting.FileSystemObject")
Set objTS = objFS.OpenTextFile("Sample.txt",1)
Dim readStr
readStr = objTS.ReadLine
msgbox readStr
