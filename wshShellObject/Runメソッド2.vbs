Option Explicit

Dim WshShell
Set WshShell = WScript.CreateObject("WScript.Shell")

'������:exe�t�@�C����
'������:�P���� �O���s��
'��O����:�P���I����҂� �O���҂����Ɏ������s
'�߂�l  :�O������I�� �P���ُ�I��
MSGBOX WshShell.Run("subst E: C:\900_Programming\VBScript\source\���K�\�[�X\wshShell�I�u�W�F�N�g",1,1)


