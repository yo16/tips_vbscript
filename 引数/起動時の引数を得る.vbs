' 引数を得る
' 2005/11/10
' 2006/04/27 memo Drag&Dropすると、フルパスが第１引数になって起動される。
'                 (0)が、第一引数。cのように自分自身のファイル名ではない。

Option Explicit

Dim objArgs, I
Set objArgs = WScript.Arguments

WScript.Echo "引数の数:" & objArgs.Count

For I = 0 to objArgs.Count - 1
	WScript.Echo "引数発見！:" & objArgs(I)
Next


