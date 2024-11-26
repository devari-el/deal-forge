Attribute VB_Name = "p_load_list_products"
Sub def_load_list_products()

Dim ws As Worksheet
Dim lastRow As Long
Dim i As Long
Dim listData() As String
Dim listBox As MSForms.listBox


Set ws = ThisWorkbook.Sheets("products")

lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).row

Set listBox = manageProducts.list_products

listBox.Clear

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
listBox.ColumnWidths = "40; 50; 125; 175; 50; 75; 60; 45; 50"
listBox.List = listData

End Sub

