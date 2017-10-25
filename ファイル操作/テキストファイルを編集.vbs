Option Explicit

'--編集ファイル名
Dim editFileName,workFileName
editFileName = "sample2.txt"
workFileName = editFileName & ".work"

Const ForReading = 1, ForWriting = 2, ForAppending = 8
Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0

Dim objFS
Set objFS = CreateObject("Scripting.FileSystemObject")

'--編集元のファイルを開く
Dim objEditFile
Set objEditFile = objFS.GetFile(editFileName)
Dim objEditTS
Set objEditTS = objEditFile.OpenAsTextStream(ForReading,TristateUseDefault)

'--編集先(Workファイル)のファイルを作成
Dim objWorkTS
Set objWorkTS = objFS.CreateTextFile(workFileName,False)

'--編集
Do Until objEditTS.AtEndOfStream
	'--この場合は[']を始めにつけている。
	objWorkTS.WriteLine "'" & objEditTS.ReadLine
Loop

'--ファイルをクローズ
objEditTS.Close
objWorkTS.Close

'--編集元のファイルを削除
objEditFile.Delete

'--編集先のファイル名を編集元のファイル名に変更
Dim objWorkFile
Set objWorkFile = objFS.GetFile(workFileName)
objWorkFile.Name = editFileName


MsgBox "終了～～～！",,"ヽ(￣▽￣)ノ"

