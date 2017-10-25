
'===事実その１===
'実行するときには、
'CScript 実行ファイル名 >書き込むファイル名
'をコマンドプロンプトで実行する
'==>ホントにできる？？(未解決 2001/02/02)

'===事実その２===
'CMD.EXE /C CScript 実行ファイル名 > 書き込むファイル名
'って書くと、stdout の操作がいらない


msgbox "標準出力.vbsの中"

WScript.Echo "stdoutを変える前。"


'CMC.EXEの場合は以下をコメントアウトしても
'ファイルに出力される
Dim stdout
Set stdout = WScript.StdOut


'-- どっちでもできる --'
'stdout.write "標準出力"
WScript.echo "標準出力"


