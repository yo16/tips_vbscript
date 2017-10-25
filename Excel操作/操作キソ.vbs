' エクセルを操作キソ
' 2007/10/10

Option Explicit


Dim objExcel 'Excelアプリ 
Dim objBook 

Dim intRowCnt 
Dim intColCnt 
Dim strLipID 
Dim strChkFlg 

Set objExcel = CreateObject("Excel.Application") 

objExcel.Visible = False 
ObjExcel.Workbooks.Open "E:\Programming\VBScript\source\練習ソース\Excel操作\test.xls" 

Set objBook = objExcel.ActiveWorkBook 

objExcel.DisplayAlerts = False 

strChkFlg = 0 


intRowCnt = 3 

strLipID = objBook.Sheets(1).Cells(intRowCnt,1) 
msgbox strLipID

objBook.Sheets(1).Cells(4,1).Value = "4-1です"
' 変更した場合は保存
objBook.Save

objBook.Close 
objExcel.Quit 

Set objBook = Nothing 
Set objExcel = Nothing 

msgbox "終了"
