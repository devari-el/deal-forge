Attribute VB_Name = "p_change_listbox_manageProducts"
Sub def_change_listbox_manageProducts()

    Dim listBox As MSForms.listBox
    Dim typeProduct As String
    Dim optService As MSForms.OptionButton
    Dim optProduct As MSForms.OptionButton
    
    ' Associa os controles
    With manageProducts
        Set listBox = .list_products
        Set optService = .opt_service
        Set optProduct = .opt_product
        
        ' Se o campo nome estiver habilitado, sair
        If .txt_name.Enabled Then Exit Sub
        
        ' Verifica se h� um item selecionado (excluindo o cabe�alho)
        If listBox.ListIndex > 0 Then
            ' Preenche os campos com base na sele��o
            .txt_code.Value = listBox.List(listBox.ListIndex, 0)
            typeProduct = listBox.List(listBox.ListIndex, 1)
            
            ' Define o tipo de produto/servi�o
            optService.Value = (typeProduct = "Servi�o")
            optProduct.Value = (typeProduct = "Produto")
            
            ' Preenche os outros campos
            .txt_name.Value = listBox.List(listBox.ListIndex, 2)
            .txt_specs.Value = listBox.List(listBox.ListIndex, 3)
            .txt_brand.Value = listBox.List(listBox.ListIndex, 4)
            .txt_supplier.Value = listBox.List(listBox.ListIndex, 5)
            
            ' Convers�o segura para valores num�ricos no padr�o americano
            If IsNumeric(listBox.List(listBox.ListIndex, 6)) Then
                .txt_weight.Value = Replace(CStr(CDbl(listBox.List(listBox.ListIndex, 6))), ",", ".")
            Else
                .txt_weight.Value = "0.00"
            End If
            
            If IsNumeric(listBox.List(listBox.ListIndex, 7)) Then
                .txt_price.Value = Replace(CStr(CDbl(listBox.List(listBox.ListIndex, 7))), ",", ".")
            Else
                .txt_price.Value = "0.00"
            End If
            
            .txt_invoice.Value = listBox.List(listBox.ListIndex, 8)
        Else
            ' Limpa os campos se nenhum item v�lido estiver selecionado
            .txt_code.Value = ""
            .txt_name.Value = ""
            .txt_specs.Value = ""
            .txt_brand.Value = ""
            .txt_supplier.Value = ""
            .txt_weight.Value = ""
            .txt_price.Value = ""
            .txt_invoice.Value = ""
            optService.Value = False
            optProduct.Value = False
        End If
    End With
    
End Sub

