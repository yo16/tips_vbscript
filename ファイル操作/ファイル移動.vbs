Option Explicit
' 移動
' 2010/06/23


' MoveFileを使う



Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")
fso.MoveFile "rename_1.txt", "rename_2.txt"

' Move先のファイルがある場合は、エラーになる

' フォルダがない場合はどうなるんだろうか。
' → 実行時エラーが出る

