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
	' ベース点出力
	objTS.WriteLine baseX & "," & baseY & "," & baseZ
	
	' 辺の長さを計算
	Dim edgeLength
	edgeLength = topLineLen * ( 0.5 ^ depth)
	
	' ２点目を計算
	Dim topX, topY, topZ
	topX = baseX + edgeLength
	topY = baseY + edgeLength
	topZ = baseZ + edgeLength
	
	' ２点目を出力
	objTS.WriteLine topX & "," & topY & "," & topZ
	
	
	' 子の点
	If ( depth < MAX_DEPTH ) Then
		' X軸の方向
		morimori baseX+edgeLength, baseY, baseZ, topLineLen, depth+1
		' Y軸の方向
		morimori baseX, baseY+edgeLength, baseZ, topLineLen, depth+1
		' Z軸の方向
		morimori baseX, baseY, baseZ+edgeLength, topLineLen, depth+1
	End If
	
	
End Sub
