' �t�@�C��������ď���
' 2016/5/12 y.ikeda

Option Explicit


Dim objFs
Set objFs = CreateObject("Scripting.FileSystemObject")

Dim fileName
fileName = "�t�@�C���폜.txt"


' �����F�쐬
Dim objTs
Set objTs = objFs.CreateTextFile(fileName)
objTs.WriteLine "������^��"
objTs.Close
Set objTs = Nothing


MsgBox fileName & "�������܂��I"


' �t�@�C�����폜����
objFs.DeleteFile fileName


MsgBox "end"
