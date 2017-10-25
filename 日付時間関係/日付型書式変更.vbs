Option Explicit

' 用意されている固定フォーマット
MsgBox	" :"&Now & vbCrLf &_
		"0:"&FormatDateTime(Now,0) & vbCrLf &_
		"1:"&FormatDateTime(Now,1) & vbCrLf &_
		"2:"&FormatDateTime(Now,2) & vbCrLf &_
		"3:"&FormatDateTime(Now,3) & vbCrLf &_
		"4:"&FormatDateTime(Now,4)
'  :2015/07/30 11:12:34
' 0:2015/07/30 11:12:34
' 1:2015年7月30日
' 2:2015/07/30
' 3:11:12:34
' 4:11:12


' カスタマイズしたいとき
' VB、VBAではFormat関数がやってくれるが
' VBSには実装されていない。
' そのため、Year、Month、Day、Hour、Minute、Secondを使って
' 自作する必要がある。
MsgBox Year(Now) & "-" & Month(Now) & "-" & Day(Now) & "*" & _
	Hour(Now) & ":" & Minute(Now) & ":" & Second(Now)
'2015-7-30*11:22:55

' １桁の場合は前ゼロが必要
' Right("0" & Month(Now), 2)のように、前ゼロをくっつけた上で
' 右から２文字採用する方法が綺麗でいい感じ。

