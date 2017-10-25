
Set WshShell = WScript.CreateObject ("WScript.Shell")



'↓これだと、ファイルに出力されない。画面にも出ない。
'intErrCode=WshShell.Run("cscript.exe 標準出力.vbs >abc.txt",0,True)
intErrCode=WshShell.Run("cmd /c cscript 標準出力.vbs >abc.txt",0,True)



'-- どっちでもできる --'
'stdout.write "aaa"
WScript.echo "aaa"


