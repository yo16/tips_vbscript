Option Explicit

' �t�H���_���R�s�[
' 2016/1/22 y.ikeda

Dim objFs
Set objFs = WScript.CreateObject("Scripting.FileSystemObject")


' from
' to
' �㏑����
objFs.CopyFolder "a", "a_copy", True

' �㏑��False�Ŋ��ɂ���ꍇ�́A���Lmsg���o��
' ---------------------------
' Windows Script Host
' ---------------------------
' �X�N���v�g:	C:\zProgramming\VBScript\source\���K�\�[�X\�h���C�u�E�t�H���_\�t�H���_���R�s�[.vbs
' �s:	13
' ����:	1
' �G���[:	���ɓ����̃t�@�C�������݂��Ă��܂��B
' �R�[�h:	800A003A
' �\�[�X: 	Microsoft VBScript ���s���G���[
' 
' ---------------------------
' OK   
' ---------------------------



msgbox "end"


