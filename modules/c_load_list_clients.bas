Attribute VB_Name = "c_load_list_clients"
Sub def_load_list_clients()

Dim ws As Worksheet
Dim lastRow As Long
Dim i As Long
Dim listData() As String
Dim listBox As MSForms.listBox

Set ws = ThisWorkbook.Sheets("clients")

lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).row

Set listBox = manageClients.list_clients

listBox.Clear

ReDim listData(1 To lastRow, 1 To 11)

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
    listData(i, 10) = ws.Cells(i, "J").Value
    listData(i, 11) = ws.Cells(i, "K").Value

Next i

listBox.ColumnCount = 11
listBox.ColumnWidths = "90; 95; 150; 44; 54; 54; 74; 29; 75; 64; 100"
listBox.List = listData

End Sub
