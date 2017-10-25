rem http://www.kanaya440.com/contents/tips/vbs/006.html
rem 2017/9/12

Dim objXML    AS Object
Dim strXMLDoc AS String
Dim intRet    AS Integer
Dim strURL    AS String
Dim strKey    AS String
 
strURL = "xxxx.co.jp"
strKey = "id=123&pass=abc"
 
Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP")
objXML.open "POST", strURL, False
objXML.setRequestHeader "Content-Type", " application/x-www-form-urlencoded"
objXML.setRequestHeader "Content-Length", "length"
objXML.send strKey
strXMLDoc = objXML.responseText
intRet = objXML.status
Set objXML = Nothing
