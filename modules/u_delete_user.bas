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
        
            ' Limpa o conteúdo da linha
            ws.Cells(i, 1).ClearContents
            ws.Cells(i, 2).ClearContents
            ws.Cells(i, 3).ClearContents
            ws.Cells(i, 4).ClearContents
            
            ' Define o intervalo abaixo da linha excluída até o último dado preenchido
            If i < lastRow Then
                ' Limita a movimentação até a coluna D (4)
                Set nextRange = ws.range(ws.Cells(i + 1, 1), ws.Cells(lastRow, 4))
                
                ' Copia o intervalo para a linha acima
                nextRange.Copy Destination:=ws.range(ws.Cells(i, 1), ws.Cells(i, 4))
                
                ' Limpa a última linha (agora duplicada após o deslocamento)
                ws.Rows(lastRow).ClearContents
            End If

            ' Atualiza a lista de usuários
            Call def_load_list_users

            Exit For

        ElseIf deletedUserName = ws.Cells(i, 6) Then

            MsgBox "Não é possível excluir o usuário atual!", vbExclamation, "DEAL FORGE"
            Exit For

        End If

    Next i

End Sub

