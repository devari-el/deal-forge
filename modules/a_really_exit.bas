Attribute VB_Name = "a_really_exit"
Sub def_really_exit()

    Dim answer As VbMsgBoxResult
    Dim ws As Worksheet
    
    Set ws = ThisWorkbook.Sheets("users")
    
    answer = MsgBox("Deseja realmente sair?", vbYesNo + vbQuestion, "DEAL FORGE")
    
    If answer = vbYes Then
    
        ' Fecha o formulário atual, e zera sessão de usuário
        ws.Cells(2, 6) = ""
        ws.Cells(2, 7) = ""
        ws.Cells(2, 8) = ""
        
        End
    End If
    
End Sub
