Attribute VB_Name = "nd_add_product_deal"
Sub def_add_product_deal()

Dim code As String
Dim listProducts As MSForms.listBox
Dim listDeal As MSForms.listBox
Dim wsProducts As Worksheet
Dim wsLayout As Worksheet
Dim lastRow As Long
Dim productName As String
Dim listData() As String
Dim i As Long
Dim j As Long
Dim k As Integer

Set wsLayout = ThisWorkbook.Sheets("layout")
Set wsProducts = ThisWorkbook.Sheets("products")
Set listProducts = newDeal.list_products
Set listDeal = newDeal.list_deal

'Checa se algum produto foi selecionado antes te adicionar uma quantidade
If listProducts.ListIndex = -1 Then
    MsgBox "Selecione um produto da lista", vbExclamation, "DEAL FORGE"
    Exit Sub
End If
    
code = listProducts.List(listProducts.ListIndex, 0)

If wsLayout.Cells(41, 4).Value <> "" Then

    MsgBox "Limite de itens atingidos!", vbCritical, "DEAL FORGE"
    Exit Sub

Else

    lastRow = wsProducts.Cells(wsProducts.Rows.count, "A").End(xlUp).row
    
    For i = 2 To lastRow
    
        If code = wsProducts.Cells(i, 1) Then
        
        Exit For
        
        End If
    
    Next i
    
    If listProducts.ListIndex = 0 And listProducts.List(0, 0) = "COD" Then
    
        Exit Sub
        
    End If
    
    If newDeal.txt_qtd.Value = "" Then
    
        MsgBox "Insira a quantidade!", vbCritical, "DEAL FORGE"
        Exit Sub
        
    End If
    
    
    'COMEÇA A PREENCHER---------------------------------------------------------
    
    For j = 15 To 45
    
    
        If wsLayout.Cells(j, 4).Value = "" Then
        
            'PREENCHER CÓD
            
            If wsProducts.Cells(i, 2).Value = "Produto" Then
            
                wsLayout.Cells(j, 2).Value = wsProducts.Cells(i, 9).Value
                
            Else
            
                wsLayout.Cells(j, 2).Value = ""
            
            End If
            '-------------------------------------------------------------------
            'PREENCHER QTD
            
            If IsNumeric(newDeal.txt_qtd.Value) Then
            
                wsLayout.Cells(j, 3).Value = CInt(newDeal.txt_qtd.Value)
                
            Else
            
                MsgBox "Por favor, insira um número inteiro válido.", vbExclamation, "DEAL FORGE"
                newDeal.txt_qtd.SetFocus
                
            End If
            '-------------------------------------------------------------------
            'PREENCHER PRODUTO
            
            productName = CStr(wsProducts.Cells(i, 3).Value & " " & wsProducts.Cells(i, 4).Value)
            wsLayout.Cells(j, 4).Value = productName
                  
            '-------------------------------------------------------------------
            'PREENCHER VALOR
                  
            wsLayout.Cells(j, 13).Value = CDbl(wsProducts.Cells(i, 8).Value)
                  
                  
        Exit For
        
        End If
    
    Next j
    
    Call def_update_listDeal
    newDeal.txt_price.Value = wsLayout.Cells(44, 10).Value
    
End If

End Sub

