Option Explicit

' �t�H���_���ړ�
' 2018/3/16 yo16

Dim objFs
Set objFs = WScript.CreateObject("Scripting.FileSystemObject")

' movetest�Ƃ����t�H���_���Aa�Ƃ����t�H���_�ȉ��Ɉړ�����
objFs.MoveFolder ".\movetest", ".\a\"

Set objFs = Nothing
