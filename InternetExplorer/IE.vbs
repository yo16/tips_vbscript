	Dim myPage,objIE
	
	myPage = InputBox("�P�F����Ӂ[" & vbCR &_
					  "�Q�F����Ӂ[���[��" & vbCR &_
 					  "�R�F�C���t�H�V�[�N")

	Set objIE = Wscript.Createobject("InternetExplorer.Application")
	Select Case myPage
		Case 1
			objIE.Navigate2  "http://www.yahoo.co.jp"
		Case 2
			objIE.Navigate2  "http://jp.f1.mail.yahoo.co.jp/ym/Login?YY=8252"
		Case 3
			objIE.Navigate2  "http://infoseek.co.jp/"
	End Select
	objIE.Visible = TRUE
