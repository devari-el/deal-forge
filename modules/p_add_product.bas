Attribute VB_Name = "p_add_product"
Sub def_add_product()

    Dim typeProduct As String
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim nextLine As Long
    Dim optConfirm As Boolean
    Dim cell As range
    Dim codeExists As Boolean

    If manageProducts.btn_modify.BackColor = RGB(0, 176, 80) Then
        MsgBox "Saia do modo Alterar Produto antes de executar esta tarefa!", vbCritical, "DEAL FORGE"
    Else
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
        
            ' Limpar os campos de entrada
            manageProducts.txt_code = ""
            manageProducts.txt_name = ""
            manageProducts.txt_specs = ""
            manageProducts.txt_brand = ""
            manageProducts.txt_supplier = ""
            manageProducts.txt_weight = ""
            manageProducts.txt_price = ""
            manageProducts.txt_invoice = ""
            manageProducts.opt_service = False
            manageProducts.opt_product = False

            manageProducts.btn_add.BackColor = RGB(0, 176, 80)
            manageProducts.btn_home.BackColor = RGB(255, 0, 0)
            manageProducts.btn_home.Caption = "CANCEL"
        Else
            Set ws = ThisWorkbook.Sheets("Products")

            ' Determinar o tipo
            If manageProducts.opt_service.Value Then
                typeProduct = "Serviço"
                optConfirm = True
            ElseIf manageProducts.opt_product.Value Then
                typeProduct = "Produto"
                optConfirm = True
            Else
                optConfirm = False
            End If

            ' Verificar campos obrigatórios
            If manageProducts.txt_code.Value = "" _
            Or manageProducts.txt_name.Value = "" _
            Or manageProducts.txt_specs.Value = "" _
            Or manageProducts.txt_brand.Value = "" _
            Or manageProducts.txt_supplier.Value = "" _
            Or manageProducts.txt_weight.Value = "" _
            Or manageProducts.txt_price.Value = "" _
            Or manageProducts.txt_invoice.Value = "" _
            Or Not optConfirm Then
                MsgBox "Preencha todos os campos corretamente!", vbCritical, "DEAL FORGE"
            Else
                ' Verificar se já existe um produto com o mesmo código
                codeExists = False
                lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).row
                
                For Each cell In ws.range("A2:A" & lastRow)
                    If cell.Value = manageProducts.txt_code.Value Then
                        codeExists = True
                        Exit For
                    End If
                Next cell

                If codeExists Then
                    MsgBox "Já existe um produto com este código!", vbExclamation, "DEAL FORGE"
                Else
                    ' Obter a próxima linha disponível
                    nextLine = lastRow + 1
                    
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

                    ' Adicionar o novo produto
                    ws.Cells(nextLine, 1).Value = manageProducts.txt_code.Value
                    ws.Cells(nextLine, 2).Value = typeProduct
                    ws.Cells(nextLine, 3).Value = manageProducts.txt_name.Value
                    ws.Cells(nextLine, 4).Value = manageProducts.txt_specs.Value
                    ws.Cells(nextLine, 5).Value = manageProducts.txt_brand.Value
                    ws.Cells(nextLine, 6).Value = manageProducts.txt_supplier.Value
                    ws.Cells(nextLine, 7).Value = CDbl(manageProducts.txt_weight.Value)
                    ws.Cells(nextLine, 8).Value = CDbl(manageProducts.txt_price.Value)
                    ws.Cells(nextLine, 9).Value = manageProducts.txt_invoice.Value

                    ' Desabilitar os campos após a adição
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

                    ' Atualizar a lista de produtos
                    Call def_load_list_products

                    manageProducts.btn_add.BackColor = RGB(25, 86, 180)
                    manageProducts.btn_home.BackColor = RGB(25, 86, 180)
                    manageProducts.btn_home.Caption = "HOME"
                End If
            End If
        End If
    End If
End Sub

