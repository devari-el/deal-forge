Attribute VB_Name = "p_delete_product"
Sub def_delete_product()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim deletedProductName As String

    Set ws = ThisWorkbook.Sheets("products")
    deletedProductName = manageProducts.txt_name.Value ' Nome do produto a ser excluído

    lastRow = ws.Cells(ws.Rows.count, "C").End(xlUp).row ' Última linha preenchida

    For i = 2 To lastRow ' Percorre as linhas a partir da segunda linha (supondo que a primeira seja cabeçalho)
        If ws.Cells(i, 3).Value = deletedProductName Then ' Compara o nome do produto
            ws.Rows(i).Delete ' Deleta a linha inteira
            
            ' Atualiza a lista de produtos (presumindo que é necessário após a exclusão)
            Call def_load_list_products
            Exit For ' Sai do loop após excluir o produto encontrado
        End If
    Next i
End Sub

