Option Explicit

' �p�ӂ���Ă���Œ�t�H�[�}�b�g
MsgBox	" :"&Now & vbCrLf &_
		"0:"&FormatDateTime(Now,0) & vbCrLf &_
		"1:"&FormatDateTime(Now,1) & vbCrLf &_
		"2:"&FormatDateTime(Now,2) & vbCrLf &_
		"3:"&FormatDateTime(Now,3) & vbCrLf &_
		"4:"&FormatDateTime(Now,4)
'  :2015/07/30 11:12:34
' 0:2015/07/30 11:12:34
' 1:2015�N7��30��
' 2:2015/07/30
' 3:11:12:34
' 4:11:12


' �J�X�^�}�C�Y�������Ƃ�
' VB�AVBA�ł�Format�֐�������Ă���邪
' VBS�ɂ͎�������Ă��Ȃ��B
' ���̂��߁AYear�AMonth�ADay�AHour�AMinute�ASecond���g����
' ���삷��K�v������B
MsgBox Year(Now) & "-" & Month(Now) & "-" & Day(Now) & "*" & _
	Hour(Now) & ":" & Minute(Now) & ":" & Second(Now)
'2015-7-30*11:22:55

' �P���̏ꍇ�͑O�[�����K�v
' Right("0" & Month(Now), 2)�̂悤�ɁA�O�[���������������
' �E����Q�����̗p������@���Y��ł��������B

