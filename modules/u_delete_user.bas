Attribute VB_Name = "u_delete_user"
Sub def_delete_user()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim deletedUserName As String
    Dim nextRange As range

    Set ws = ThisWorkbook.Sheets("users")

    deletedUserName = manageUsers.txt_name.Value
    lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).row

    For i = 2 To lastRow

        If deletedUserName = ws.Cells(i, 1) Then
        
            ' Limpa o conte�do da linha
            ws.Cells(i, 1).ClearContents
            ws.Cells(i, 2).ClearContents
            ws.Cells(i, 3).ClearContents
            ws.Cells(i, 4).ClearContents
            
            ' Define o intervalo abaixo da linha exclu�da at� o �ltimo dado preenchido
            If i < lastRow Then
                ' Limita a movimenta��o at� a coluna D (4)
                Set nextRange = ws.range(ws.Cells(i + 1, 1), ws.Cells(lastRow, 4))
                
                ' Copia o intervalo para a linha acima
                nextRange.Copy Destination:=ws.range(ws.Cells(i, 1), ws.Cells(i, 4))
                
                ' Limpa a �ltima linha (agora duplicada ap�s o deslocamento)
                ws.Rows(lastRow).ClearContents
            End If

            ' Atualiza a lista de usu�rios
            Call def_load_list_users

            Exit For

        ElseIf deletedUserName = ws.Cells(i, 6) Then

            MsgBox "N�o � poss�vel excluir o usu�rio atual!", vbExclamation, "DEAL FORGE"
            Exit For

        End If

    Next i

End Sub

