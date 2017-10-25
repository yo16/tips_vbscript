Option Explicit


Function countHeadTab(str) {
	Dim rtnCount = 0
	
	Dim regPattern = "^\t+"
	Dim regEx
	Set regEx = New RegExp
	regEx.Pattern = regPattern
	Dim Matches
	Set Matches = regEx.Execute(str)
	
