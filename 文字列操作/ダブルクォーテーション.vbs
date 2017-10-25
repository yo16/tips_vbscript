Option Explicit


'２つ連続で書くと１つの"となる
msgbox """"

'変数での操作
Dim testStr
testStr = """abc"""
msgbox testStr

'ファイルの操作
Dim objFS,objTS
Set objFS = WScript.CreateObject("Scripting.FileSystemObject")
Set objTS = objFS.OpenTextFile("Sample.txt",1)
Dim readStr
readStr = objTS.ReadLine
msgbox readStr
