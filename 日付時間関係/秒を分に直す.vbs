Option Explicit

Dim byou
byou = 150

msgbox ByouToFun(byou)


'60‚ÅŠ„‚Á‚ÄØ‚èÌ‚Ä‚é‚¾‚¯‚¾‚æH
Function ByouToFun(pTime)
	Dim rtnByou
	rtnByou = Int(pTime)

	ByouToFun = rtnByou \ 60
End Function


