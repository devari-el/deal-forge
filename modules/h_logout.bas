Attribute VB_Name = "h_logout"
Sub def_logout()

    Dim answer As VbMsgBoxResult
    Dim ws As Worksheet
    
    Set ws = ThisWorkbook.Sheets("users")
    
    answer = MsgBox("Deseja realmente sair?", vbYesNo + vbQuestion, "DEAL FORGE")
    
    If answer = vbYes Then
    
        ' Fecha o formulário atual, zera sessão de usuário
        ws.Cells(2, 6) = ""
        ws.Cells(2, 7) = ""
        ws.Cells(2, 8) = ""
        
        Unload home
        login.Show
    
    End If

End Sub
