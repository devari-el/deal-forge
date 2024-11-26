Attribute VB_Name = "nd_uptade_listDeal"
Sub def_update_listDeal()

    Dim ws As Worksheet
    Dim i As Long
    Dim listData() As Variant ' Usado para armazenar os dados da planilha
    Dim listBox As MSForms.listBox
    Dim lastRow As Long
    Dim itemValue As Variant ' Variável para verificar o valor

    ' Definir as referências
    Set ws = ThisWorkbook.Sheets("layout")
    Set listBox = newDeal.list_deal

    ' Calcular a última linha preenchida no intervalo
    lastRow = ws.range("D41").End(xlUp).row
    If lastRow < 15 Then lastRow = 15

    ' Redimensionar o array para acomodar os dados
    ReDim listData(1 To lastRow - 14, 1 To 4) ' Linhas: 15 até lastRow

    ' Preencher o array com os dados da planilha (ignorando zeros)
    Dim rowIdx As Integer
    rowIdx = 1
    
    For i = 15 To lastRow
        ' Ignorar linhas onde a coluna D está vazia ou zero
        If ws.Cells(i, 4).Value <> "" And ws.Cells(i, 4).Value <> 0 Then
            listData(rowIdx, 1) = ws.Cells(i, "C").Value  ' Quantidade
            listData(rowIdx, 2) = ws.Cells(i, "D").Value  ' Produto
            listData(rowIdx, 3) = ws.Cells(i, "H").Value  ' Valor Unitário
            listData(rowIdx, 4) = ws.Cells(i, "J").Value  ' Valor Total

            rowIdx = rowIdx + 1
        End If
    Next i

    ' Verificar se há dados no array antes de redimensionar
    If rowIdx > 1 Then
        ' Redimensionar o array para remover linhas vazias extras
        ReDim Preserve listData(1 To rowIdx - 1, 1 To 4)

        ' Atualizar a ListBox com os dados combinados
        listBox.ColumnCount = 4
        listBox.ColumnWidths = "15; 175; 40; 40"
        listBox.List = listData
    Else
        ' Se não houver dados, limpar a ListBox
        listBox.Clear
    End If

End Sub

