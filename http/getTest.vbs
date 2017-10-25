Dim objXML    AS Object
Dim strXMLDoc AS String
Dim intRet    AS Integer
Dim strURL    AS String
Dim strKey    AS String

strURL = "is-tcc-mail2pcb/SupportLog/Graph/cgi/ListupNoMail.pl"
'strKey = "id=123&pass=abc"

Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP")
objXML.open "GET", strURL, False
objXML.setRequestHeader "Content-Type", " application/x-www-form-urlencoded"
objXML.setRequestHeader "Content-Length", "length"
objXML.send strKey
strXMLDoc = objXML.responseText
intRet = objXML.status
Set objXML = Nothing
