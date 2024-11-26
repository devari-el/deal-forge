Attribute VB_Name = "p_modify_products"
Sub def_modify_product()

    Dim ws As Worksheet
    Dim selectedIndex As Long
    Dim rowToEdit As Long
    Dim typeProduct As String
    Dim existingRow As Long
    Dim i As Long
    Dim code As String
    Dim codeAlreadyExists As Boolean

    Set ws = ThisWorkbook.Sheets("products")

    ' Verificar se há campos vazios
    If manageProducts.txt_code = "" _
    Or manageProducts.txt_name = "" _
    Or manageProducts.txt_specs = "" _
    Or manageProducts.txt_brand = "" _
    Or manageProducts.txt_supplier = "" _
    Or manageProducts.txt_weight = "" _
    Or manageProducts.txt_price = "" _
    Or manageProducts.txt_invoice = "" _
    Or (Not manageProducts.opt_service.Value And Not manageProducts.opt_product.Value) Then

        MsgBox "Existem campos vazios!", vbInformation, "DEAL FORGE"
        Exit Sub
    End If

    ' Ativar modo de edição
    If manageProducts.txt_name.Enabled = False Then
        ' Habilitar os campos de entrada
        manageProducts.txt_code.Enabled = True
        manageProducts.txt_name.Enabled = True
        manageProducts.txt_specs.Enabled = True
        manageProducts.txt_brand.Enabled = True
        manageProducts.txt_supplier.Enabled = True
        manageProducts.txt_weight.Enabled = True
        manageProducts.txt_price.Enabled = True
        manageProducts.txt_invoice.Enabled = True
        manageProducts.opt_service.Enabled = True
        manageProducts.opt_product.Enabled = True

        manageProducts.btn_modify.BackColor = RGB(0, 176, 80)
        manageProducts.btn_home.BackColor = RGB(255, 0, 0)
        manageProducts.btn_home.Caption = "CANCEL"

    Else
        ' Obter índice do item selecionado
        selectedIndex = manageProducts.list_products.ListIndex

        ' Verificar se um item foi selecionado
        If selectedIndex = -1 Then
            MsgBox "Selecione um produto para modificar!", vbExclamation, "DEAL FORGE"
            Exit Sub
        End If

        ' Obter a linha correspondente na planilha
        rowToEdit = selectedIndex + 1 ' Supondo que o cabeçalho está na linha 1

        ' Verificar se o código já existe na coluna A
        code = manageProducts.txt_code.Value
        codeAlreadyExists = False

        For i = 2 To ws.Cells(ws.Rows.count, "A").End(xlUp).row
            If ws.Cells(i, 1).Value = code And i <> rowToEdit Then
                codeAlreadyExists = True
                Exit For
            End If
        Next i

        If codeAlreadyExists Then
            MsgBox "Este código de produto já existe!", vbCritical, "DEAL FORGE"
            Exit Sub
        End If

        ' Determinar o tipo de produto
        If manageProducts.opt_service.Value Then
            typeProduct = "Serviço"
        Else
            typeProduct = "Produto"
        End If

        ' Verificar se txt_weight é numérico
        If Not IsNumeric(manageProducts.txt_weight.Value) Then
            MsgBox "O valor do peso deve ser numérico.", vbExclamation, "DEAL FORGE"
            Exit Sub
        End If
        
        ' Verificar se txt_price é numérico
        If Not IsNumeric(manageProducts.txt_price.Value) Then
            MsgBox "O valor do preço deve ser numérico.", vbExclamation, "DEAL FORGE"
            Exit Sub
        End If

        ' Atualizar os dados na planilha
        ws.Cells(rowToEdit, 1).Value = manageProducts.txt_code.Value
        ws.Cells(rowToEdit, 2).Value = typeProduct
        ws.Cells(rowToEdit, 3).Value = manageProducts.txt_name.Value
        ws.Cells(rowToEdit, 4).Value = manageProducts.txt_specs.Value
        ws.Cells(rowToEdit, 5).Value = manageProducts.txt_brand.Value
        ws.Cells(rowToEdit, 6).Value = manageProducts.txt_supplier.Value
        ws.Cells(rowToEdit, 7).Value = CDbl(manageProducts.txt_weight.Value)
        ws.Cells(rowToEdit, 8).Value = CDbl(manageProducts.txt_price.Value)
        ws.Cells(rowToEdit, 9).Value = manageProducts.txt_invoice.Value

        ' Atualizar a ListBox
        manageProducts.list_products.List(selectedIndex, 0) = manageProducts.txt_code.Value
        manageProducts.list_products.List(selectedIndex, 1) = typeProduct
        manageProducts.list_products.List(selectedIndex, 2) = manageProducts.txt_name.Value
        manageProducts.list_products.List(selectedIndex, 3) = manageProducts.txt_specs.Value
        manageProducts.list_products.List(selectedIndex, 4) = manageProducts.txt_brand.Value
        manageProducts.list_products.List(selectedIndex, 5) = manageProducts.txt_supplier.Value
        manageProducts.list_products.List(selectedIndex, 6) = CDbl(manageProducts.txt_weight.Value)
        manageProducts.list_products.List(selectedIndex, 7) = CDbl(manageProducts.txt_price.Value)
        manageProducts.list_products.List(selectedIndex, 8) = manageProducts.txt_invoice.Value

        ' Desabilitar os campos após a modificação
        manageProducts.txt_code.Enabled = False
        manageProducts.txt_name.Enabled = False
        manageProducts.txt_specs.Enabled = False
        manageProducts.txt_brand.Enabled = False
        manageProducts.txt_supplier.Enabled = False
        manageProducts.txt_weight.Enabled = False
        manageProducts.txt_price.Enabled = False
        manageProducts.txt_invoice.Enabled = False
        manageProducts.opt_service.Enabled = False
        manageProducts.opt_product.Enabled = False

        manageProducts.btn_modify.BackColor = RGB(25, 86, 180)
        manageProducts.btn_home.BackColor = RGB(25, 86, 180)
        manageProducts.btn_home.Caption = "HOME"

        Call def_load_list_products
    End If

End Sub


