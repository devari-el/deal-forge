Attribute VB_Name = "md_controllers"
Sub def_load_comb_status()

Dim cmbStatus As MSForms.ComboBox
Set cmbStatus = manageDeals.comb_status

cmbStatus.AddItem "Emitida"
cmbStatus.AddItem "Enviada"
cmbStatus.AddItem "Aprovada"
cmbStatus.AddItem "Faturada"
cmbStatus.AddItem "Recebida"

End Sub
Sub def_load_list_deals()

Dim ws As Worksheet
Dim listBox As MSForms.listBox
Dim listData() As String
Dim lastRow As Long
Dim i As Long

Set ws = ThisWorkbook.Sheets("deals")
Set listBox = manageDeals.list_deals
listBox.Clear

lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

ReDim listData(1 To lastRow, 1 To 9)

For i = 1 To lastRow

    listData(i, 1) = ws.Cells(i, "A").Value
    listData(i, 2) = ws.Cells(i, "B").Value
    listData(i, 3) = ws.Cells(i, "C").Value
    listData(i, 4) = ws.Cells(i, "D").Value
    listData(i, 5) = ws.Cells(i, "E").Value
    listData(i, 6) = ws.Cells(i, "F").Value
    listData(i, 7) = ws.Cells(i, "G").Value
    listData(i, 8) = ws.Cells(i, "H").Value
    listData(i, 9) = ws.Cells(i, "I").Value
    
Next i

listBox.ColumnCount = 9
listBox.list = listData

End Sub
Sub def_list_deals_change()

Dim listDeals As MSForms.listBox
Dim cmbStatus As MSForms.ComboBox
Set listDeals = manageDeals.list_deals
Set cmbStatus = manageDeals.comb_status

If listDeals.ListIndex > 0 Then
    cmbStatus = listDeals.list(listDeals.ListIndex, 8)
Else
    cmbStatus = ""
End If

End Sub
Sub def_update_status()

Dim ws As Worksheet
Dim listDeals As MSForms.listBox
Dim cmbStatus As MSForms.ComboBox
Dim selectedDealIndex As Long

Set ws = ThisWorkbook.Sheets("deals")
Set listDeals = manageDeals.list_deals
Set cmbStatus = manageDeals.comb_status
selectedDealIndex = listDeals.ListIndex

If selectedDealIndex < 1 Then
    MsgBox ("Escolha um orçamento para atualizar o status"), vbExclamation, "DEAL FORGE"
Else
    Dim lastRow As Long
    Dim newStatus As String
    Dim selectedDealId As String
    
    newStatus = cmbStatus.Value
    lastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row
    selectedDealId = listDeals.list(selectedDealIndex, 0)
    
    If newStatus = "Emitida" Or _
       newStatus = "Enviada" Or _
       newStatus = "Aprovada" Or _
       newStatus = "Faturada" Or _
       newStatus = "Recebida" Then
    Else
        MsgBox "Status inválido!", vbExclamation
        Exit Sub
    End If
    
    For i = 1 To lastRow
        If ws.Cells(i, "A").Value = selectedDealId Then
            ws.Cells(i, "I").Value = newStatus
            MsgBox ("Status atualizado com sucesso"), vbInformation, "DEAL FORGE"
            Call def_load_list_deals 'Atualizar a lista
        End If
    Next i
    
End If

End Sub



