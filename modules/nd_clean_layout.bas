Attribute VB_Name = "nd_clean_layout"
Sub def_clean_layout()

Dim ws As Worksheet
Dim i As Integer

Set ws = ThisWorkbook.Sheets("layout")

For i = 15 To 41
    
    ws.Cells(i, 2).Value = ""
    ws.Cells(i, 3).Value = ""
    ws.Cells(i, 4).Value = ""
    ws.Cells(i, 13).Value = ""

Next i

ws.Cells(2, 5).Value = ""
ws.Cells(4, 5).Value = ""
ws.Cells(5, 5).Value = ""
ws.Cells(6, 5).Value = ""
ws.Cells(4, 10).Value = ""
ws.Cells(9, 10).Value = ""
ws.Cells(10, 10).Value = ""
ws.Cells(43, 10).Value = ""
ws.Cells(46, 10).Value = ""
ws.Cells(47, 10).Value = ""
ws.Cells(43, 4).Value = ""
ws.Cells(44, 4).Value = ""
ws.Cells(47, 3).Value = ""

For i = 8 To 12

    ws.Cells(i, 3).Value = ""

Next i

newDeal.list_deal.Clear
newDeal.txt_price.Value = ""

End Sub
