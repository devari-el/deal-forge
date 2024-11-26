Attribute VB_Name = "p_delete_product"
Sub def_delete_product()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim deletedProductName As String

    Set ws = ThisWorkbook.Sheets("products")
    deletedProductName = manageProducts.txt_name.Value ' Nome do produto a ser exclu�do

    lastRow = ws.Cells(ws.Rows.count, "C").End(xlUp).row ' �ltima linha preenchida

    For i = 2 To lastRow ' Percorre as linhas a partir da segunda linha (supondo que a primeira seja cabe�alho)
        If ws.Cells(i, 3).Value = deletedProductName Then ' Compara o nome do produto
            ws.Rows(i).Delete ' Deleta a linha inteira
            
            ' Atualiza a lista de produtos (presumindo que � necess�rio ap�s a exclus�o)
            Call def_load_list_products
            Exit For ' Sai do loop ap�s excluir o produto encontrado
        End If
    Next i
End Sub

