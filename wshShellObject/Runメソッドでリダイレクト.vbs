Option Explicit

Dim WshShell
Set WshShell = WScript.CreateObject("WScript.Shell")

'������:exe�t�@�C����
'������:�P���� �O���s��
'��O����:�P���I����҂� �O���҂����Ɏ������s
'�߂�l  :�O������I�� �P���ُ�I��
Dim runRtn
runRtn = WshShell.Run("cmd /C echo %date% %time% > date.txt",1,1)

' del�Ɠ�������

msgbox runRtn
