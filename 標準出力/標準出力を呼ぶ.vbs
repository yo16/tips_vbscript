
Set WshShell = WScript.CreateObject ("WScript.Shell")



'�����ꂾ�ƁA�t�@�C���ɏo�͂���Ȃ��B��ʂɂ��o�Ȃ��B
'intErrCode=WshShell.Run("cscript.exe �W���o��.vbs >abc.txt",0,True)
intErrCode=WshShell.Run("cmd /c cscript �W���o��.vbs >abc.txt",0,True)



'-- �ǂ����ł��ł��� --'
'stdout.write "aaa"
WScript.echo "aaa"


