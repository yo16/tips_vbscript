' �N���b�v�{�[�h�̃e�L�X�g���擾����
Function GetClipboardText()
    Dim objHTML
    Set objHTML = CreateObject("htmlfile")
    GetClipboardText = Trim(objHTML.ParentWindow.ClipboardData.GetData("text"))
End Function


msgbox GetClipboardText
