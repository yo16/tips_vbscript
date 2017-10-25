' クリップボードのテキストを取得する
Function GetClipboardText()
    Dim objHTML
    Set objHTML = CreateObject("htmlfile")
    GetClipboardText = Trim(objHTML.ParentWindow.ClipboardData.GetData("text"))
End Function


msgbox GetClipboardText
