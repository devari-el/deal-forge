Attribute VB_Name = "nd_delete_item"
Sub def_delete_item()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim deletedItem As String
    Dim found As Boolean
    Dim nextRange As range

    Set ws = ThisWorkbook.Sheets("layout")

    ' Garantir que um item foi selecionado
    If newDeal.list_deal.ListIndex = -1 Then
        MsgBox "Selecione um item para excluir.", vbExclamation, "DEAL FORGE"
        Exit Sub
    End If

    ' Capturar o valor do item selecionado (coluna 4 da planilha)
    deletedItem = Trim(newDeal.list_deal.List(newDeal.list_deal.ListIndex, 1))

    ' Encontrar a última linha preenchida no intervalo
    lastRow = ws.range("D41").End(xlUp).row
    If lastRow < 15 Then lastRow = 15

    ' Inicializar a variável de controle
    found = False

    ' Percorrer as linhas para encontrar o item a ser deletado
    For i = lastRow To 15 Step -1
        ' Comparar removendo espaços e ignorando maiúsculas/minúsculas
        If StrComp(Trim(ws.Cells(i, 4).Value), deletedItem, vbTextCompare) = 0 Then
            
            ' Limpar o conteúdo da linha selecionada (colunas B, C, D e M)
            ws.Cells(i, 2).ClearContents  ' Coluna B
            ws.Cells(i, 3).ClearContents  ' Coluna C
            ws.Cells(i, 4).Value = ""  ' Coluna D
            ws.Cells(i, 13).ClearContents ' Coluna M

            found = True  ' Marcar como encontrado
            Exit For      ' Termina o loop após deletar o item
        End If
    Next i
    
    If i < lastRow Then
        ' Percorrer da linha excluída até a penúltima linha
        For j = i To lastRow - 1
            ws.Cells(j, 2).Value = ws.Cells(j + 1, 2).Value  ' Coluna B
            ws.Cells(j, 3).Value = ws.Cells(j + 1, 3).Value  ' Coluna C
            ws.Cells(j, 4).Value = ws.Cells(j + 1, 4).Value  ' Coluna D
            ws.Cells(j, 13).Value = ws.Cells(j + 1, 13).Value ' Coluna M
        Next j
        
        ' Limpar a última linha, agora duplicada após a movimentação
        ws.Cells(lastRow, 2) = ""  ' Coluna B
        ws.Cells(lastRow, 3) = ""  ' Coluna C
        ws.Cells(lastRow, 4) = ""  ' Coluna D
        ws.Cells(lastRow, 13) = "" ' Coluna M
    End If
    
    newDeal.txt_price.Value = ws.Cells(44, 10).Value
    
    Call def_update_listDeal
        
End Sub


