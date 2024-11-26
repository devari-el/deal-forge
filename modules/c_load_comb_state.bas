Attribute VB_Name = "c_load_comb_state"
Sub def_load_comb_state()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long

    Set ws = ThisWorkbook.Sheets("clients")
    lastRow = ws.Cells(ws.Rows.count, "M").End(xlUp).row

    ' Limpa os itens existentes no combobox, caso necessário
    manageClients.comb_state.Clear

    ' Loop para adicionar os valores da coluna M ao combobox
    For i = 1 To lastRow
        manageClients.comb_state.AddItem ws.Cells(i, 13).Value
    Next i

End Sub
