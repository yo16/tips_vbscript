Option Explicit

Dim WshShell
Set WshShell = WScript.CreateObject("WScript.Shell")

'������:exe�t�@�C����
'������:�P���� �O���s��
'��O����:�P���I����҂� �O���҂����Ɏ������s
'�߂�l  :�O������I�� �P���ُ�I��
'MSGBOX WshShell.Run("copy sample.txt sample_cpy.txt",1,1)		' ��copy�̓R�}���h������ł��Ȃ��炵��(������del���Q��)
MSGBOX WshShell.Run("xcopy sample.txt sample_cpy.txt",1,1)		' ��xcopy�͂ł���B


