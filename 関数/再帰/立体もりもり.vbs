Option Explicit


Dim fileName
fileName = "abc.txt"
Dim overWrite
overWrite = True

Dim objFS,objTS
Set objFS = WScript.CreateObject("Scripting.FileSystemObject")
Set objTS = objFS.CreateTextFile(fileName,overWrite)


Dim MAX_DEPTH
MAX_DEPTH = 5
'****************
morimori 0,0,0, 100, 0
'****************

objTS.Close




Sub morimori( baseX,baseY,baseZ, topLineLen, depth )
	' �x�[�X�_�o��
	objTS.WriteLine baseX & "," & baseY & "," & baseZ
	
	' �ӂ̒������v�Z
	Dim edgeLength
	edgeLength = topLineLen * ( 0.5 ^ depth)
	
	' �Q�_�ڂ��v�Z
	Dim topX, topY, topZ
	topX = baseX + edgeLength
	topY = baseY + edgeLength
	topZ = baseZ + edgeLength
	
	' �Q�_�ڂ��o��
	objTS.WriteLine topX & "," & topY & "," & topZ
	
	
	' �q�̓_
	If ( depth < MAX_DEPTH ) Then
		' X���̕���
		morimori baseX+edgeLength, baseY, baseZ, topLineLen, depth+1
		' Y���̕���
		morimori baseX, baseY+edgeLength, baseZ, topLineLen, depth+1
		' Z���̕���
		morimori baseX, baseY, baseZ+edgeLength, topLineLen, depth+1
	End If
	
	
End Sub
