Attribute VB_Name = "c_delete_client"
Sub def_delete_client()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim deletedClientName As String
    Dim rangeToMove As range

    Set ws = ThisWorkbook.Sheets("clients")
    deletedClientName = manageClients.txt_name.Value ' Nome do cliente a ser excluído

    lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).row ' Última linha preenchida

    For i = 2 To lastRow ' Percorre as linhas a partir da segunda linha (supondo que a primeira seja cabeçalho)
        If ws.Cells(i, 1).Value = deletedClientName Then ' Compara o nome do cliente
            ' Limpa os dados nas colunas relevantes (não exclui a linha inteira)
            ws.Cells(i, 1).ClearContents ' Nome
            ws.Cells(i, 2).ClearContents ' CNPJ
            ws.Cells(i, 3).ClearContents ' Rua
            ws.Cells(i, 4).ClearContents ' Número
            ws.Cells(i, 5).ClearContents ' Bairro
            ws.Cells(i, 6).ClearContents ' CEP
            ws.Cells(i, 7).ClearContents ' Cidade
            ws.Cells(i, 8).ClearContents ' Estado
            ws.Cells(i, 9).ClearContents ' Telefone
            ws.Cells(i, 10).ClearContents ' Comprador
            ws.Cells(i, 11).ClearContents ' E-mail

            ' Se não for a última linha, move as linhas abaixo para cima
            If i < lastRow Then
                ' Definir o intervalo de A até K da próxima linha até a última linha de dados
                Set rangeToMove = ws.range(ws.Cells(i + 1, 1), ws.Cells(lastRow, 11))

                ' Move as linhas para cima (sem sobrescrever dados)
                rangeToMove.Cut Destination:=ws.Cells(i, 1)

                ' Limpar as células originais que foram movidas
                ws.range(ws.Cells(lastRow, 1), ws.Cells(lastRow, 11)).ClearContents
            End If

            Exit For ' Sai do loop após excluir os dados encontrados
        End If
    Next i

    ' Atualiza a lista de clientes após a exclusão
    Call def_load_list_clients
    
    newDeal.txt_price.Value = ws.Cells(44, 10).Value
    
End Sub


