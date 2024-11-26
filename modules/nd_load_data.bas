Attribute VB_Name = "nd_load_data"
Sub def_load_data()

Dim i As Long
Dim wsData As Worksheet
Dim wsClients As Worksheet
Dim wsProducts As Worksheet
Dim lastRow As Long
Dim listData() As String
Dim listBox As MSForms.listBox

Set wsData = ThisWorkbook.Sheets("data")
Set wsClients = ThisWorkbook.Sheets("clients")


newDeal.comb_client.Clear
newDeal.comb_conditions.Clear
newDeal.comb_delivery.Clear
newDeal.comb_term.Clear

'---------------------------------------------------------------------------

newDeal.comb_delivery.AddItem wsData.Cells(5, 1).Value
newDeal.comb_delivery.AddItem wsData.Cells(6, 1).Value

'---------------------------------------------------------------------------

For i = 5 To 15

    newDeal.comb_term.AddItem wsData.Cells(i, 2).Value
    
Next i

'---------------------------------------------------------------------------

For i = 5 To 12

    newDeal.comb_conditions.AddItem wsData.Cells(i, 3).Value
    
Next i

'---------------------------------------------------------------------------

lastRow = wsClients.Cells(wsClients.Rows.count, "A").End(xlUp).row

For i = 2 To lastRow

    newDeal.comb_client.AddItem wsClients.Cells(i, 1).Value

Next i

Call def_load_list_products_nd
Call def_update_listDeal

End Sub


Sub def_load_list_products_nd()

Dim i As Long
Dim wsProducts As Worksheet
Dim lastRow As Long
Dim listData() As String
Dim listBox As MSForms.listBox

Set wsProducts = ThisWorkbook.Sheets("products")

lastRow = wsProducts.Cells(wsProducts.Rows.count, "A").End(xlUp).row

Set listBox = newDeal.list_products

listBox.Clear

ReDim listData(1 To lastRow, 1 To 9)


For i = 1 To lastRow

    listData(i, 1) = wsProducts.Cells(i, "A").Value
    listData(i, 2) = wsProducts.Cells(i, "B").Value
    listData(i, 3) = wsProducts.Cells(i, "C").Value
    listData(i, 4) = wsProducts.Cells(i, "D").Value
    listData(i, 5) = wsProducts.Cells(i, "E").Value
    listData(i, 6) = wsProducts.Cells(i, "F").Value
    listData(i, 7) = wsProducts.Cells(i, "G").Value
    listData(i, 8) = wsProducts.Cells(i, "H").Value
    listData(i, 9) = wsProducts.Cells(i, "I").Value

Next i

listBox.ColumnCount = 9
listBox.ColumnWidths = "30; 0; 125; 150; 0; 0; 0; 45; 50"
listBox.List = listData

End Sub
