Option Explicit

Dim WshShell
Set WshShell = WScript.CreateObject("WScript.Shell")

'������:exe�t�@�C����
'������:�P���� �O���s��
'��O����:�P���I����҂� �O���҂����Ɏ������s
'�߂�l  :�O������I�� �P���ُ�I��
MSGBOX WshShell.Run("C:\winnt\system32\cmd.exe",1,1)


