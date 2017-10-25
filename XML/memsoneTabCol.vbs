Option Explicit
' memsone8.0.xmlから
' テーブルと列を取り出す

' テキストファイルオープン
Dim objFS,objTS
Set objFS = WScript.CreateObject("Scripting.FileSystemObject")
Set objTS = objFS.CreateTextFile("memsoneTabCol7.0.txt",TRUE)



Dim objDOM, rtResult

Set objDOM = WScript.CreateObject("MSXML2.DOMDocument")
rtResult = objDOM.load("memsone7.0.xml")
If rtResult = True Then
	Dim obj1, obj2, obj3, obj4, obj5, obj6
	Dim tabAtr, colAtr
	Dim tabName, colName, colType
	Dim tabNo, colNo
	tabNo = 0
	' obj1
	For Each obj1 In objDOM.childNodes
		If ( ( obj1.nodeName = "DBMODEL" ) and obj1.hasChildNodes ) Then
			' obj2
			For Each obj2 In obj1.childNodes
				If ( ( obj2.nodeName = "METADATA" ) and obj2.hasChildNodes ) Then
					' obj3
					For Each obj3 In obj2.childNodes
					If ( ( obj3.nodeName = "TABLES" ) and obj3.hasChildNodes ) Then
						' obj4
						For Each obj4 In obj3.childNodes
							
							' テーブル
							If ( obj4.nodeName = "TABLE" ) Then
								tabNo = tabNo + 1
								For Each tabAtr In obj4.attributes
									If ( tabAtr.Name = "Tablename" ) Then
										tabName = tabAtr.Value
									End If
								Next
								
								' obj5
								For Each obj5 In obj4.childNodes
									If ( ( obj5.nodeName = "COLUMNS" ) and obj5.hasChildNodes ) Then
										colNo = 0
										' obj6
										For Each obj6 In obj5.childNodes
											
											' 列
											If ( obj6.nodeName = "COLUMN" ) Then
												colNo = colNo + 1
												For Each colAtr In obj6.attributes
													If ( colAtr.Name = "ColName" ) Then
														colName = colAtr.Value
													ElseIf ( colAtr.Name = "idDatatype" ) Then
														If ( colAtr.Value = 5 ) Then
															colType = "INT"
														ElseIf ( colAtr.Value = 20 ) Then
															colType = "VARCHAR"
														ElseIf ( colAtr.Value = 14 ) Then
															colType = "DATE"
														ElseIf ( colAtr.Value = 28 ) Then
															colType = "TEXT"
														ElseIf ( colAtr.Value = 11 ) Then
															colType = "REAL"
														ElseIf ( colAtr.Value = 31 ) Then
															colType = "ENUM"
														Else
															colType = colAtr.Value
														End If
													End If
												Next
												
												' ファイル出力（列名）
												objTS.WriteLine tabNo & vbTab & tabName & vbTab & colNo & vbTab & colName & vbTab & colType
											End If
										Next
									End If
								Next
							End If
						Next
					End If
					Next
				End If
			Next
		End If
	Next
End If
Set objDOM = Nothing

objTS.Close

