Attribute VB_Name = "u_load_list_users"
Sub def_load_list_users()

Dim ws As Worksheet
Dim lastRow As Long
Dim i As Long
Dim listData() As String
Dim listBox As MSForms.listBox

Set ws = ThisWorkbook.Sheets("users")

lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).row

Set listBox = manageUsers.list_users

listBox.Clear

ReDim listData(1 To lastRow, 1 To 4)

For i = 1 To lastRow

    listData(i, 1) = ws.Cells(i, "A").Value
    listData(i, 2) = ws.Cells(i, "B").Value
    listData(i, 3) = ws.Cells(i, "C").Value
    listData(i, 4) = ws.Cells(i, "D").Value

Next i

listBox.ColumnCount = 4
listBox.ColumnWidths = "125; 125; 0; 125"
listBox.List = listData

End Sub
