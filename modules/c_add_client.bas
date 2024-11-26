Attribute VB_Name = "c_add_client"
Sub def_add_client()

Dim typeProduct As String
Dim ws As Worksheet
Dim lastRow As Long
Dim nextLine As Long
Dim cell As range
Dim cnpjExists As Boolean

If manageClients.btn_modify.BackColor = RGB(0, 176, 80) Then
    
    MsgBox "Saia do modo Alterar Cliente antes de executar esta tarefa!", vbCritical, "DEAL FORGE"

Else

    If manageClients.txt_name.Enabled = False Then
    
        ' Habilitar os campos de entrada
        manageClients.txt_name.Enabled = True
        manageClients.txt_cnpj.Enabled = True
        manageClients.txt_street.Enabled = True
        manageClients.txt_number.Enabled = True
        manageClients.txt_nbhood.Enabled = True
        manageClients.txt_zipcode.Enabled = True
        manageClients.txt_city.Enabled = True
        manageClients.comb_state.Enabled = True
        manageClients.txt_phone_number.Enabled = True
        manageClients.txt_buyer.Enabled = True
        manageClients.txt_email.Enabled = True
    
        ' Limpar os campos de entrada
        manageClients.txt_name.Value = ""
        manageClients.txt_cnpj.Value = ""
        manageClients.txt_street.Value = ""
        manageClients.txt_number.Value = ""
        manageClients.txt_nbhood.Value = ""
        manageClients.txt_zipcode.Value = ""
        manageClients.txt_city.Value = ""
        manageClients.comb_state.Value = ""
        manageClients.txt_phone_number.Value = ""
        manageClients.txt_buyer.Value = ""
        manageClients.txt_email.Value = ""

        
        manageClients.btn_add.BackColor = RGB(0, 176, 80)
        manageClients.btn_home.BackColor = RGB(255, 0, 0)
        manageClients.btn_home.Caption = "CANCEL"
    
    Else
    
        Set ws = ThisWorkbook.Sheets("clients")
    
        ' Verificar campos obrigatórios
        If manageClients.txt_name.Value = "" _
        Or manageClients.txt_cnpj.Value = "" _
        Or manageClients.txt_street.Value = "" _
        Or manageClients.txt_number.Value = "" _
        Or manageClients.txt_nbhood.Value = "" _
        Or manageClients.txt_zipcode.Value = "" _
        Or manageClients.txt_city.Value = "" _
        Or manageClients.comb_state.Value = "" _
        Or manageClients.txt_phone_number.Value = "" _
        Or manageClients.txt_buyer.Value = "" _
        Or manageClients.txt_email.Value = "" Then


            MsgBox "Preencha todos os campos!", vbCritical, "DEAL FORGE"
        
        Else
            
            ' Verificar se já existe um produto com o mesmo código
            cnpjExists = False
            lastRow = ws.Cells(ws.Rows.count, "B").End(xlUp).row
            
            For Each cell In ws.range("B2:B" & lastRow) ' Verifica na coluna de usernames
            
                If cell.Value = manageClients.txt_cnpj.Value Then
                    cnpjExists = True
                    Exit For
                    
                End If
                
            Next cell
            
            If cnpjExists Then
                MsgBox "Já existe um cliente com este CNPJ!", vbExclamation, "DEAL FORGE"
            Else
                ' Adicionar o novo usuário
                nextLine = lastRow + 1
                ws.Cells(nextLine, 1).Value = manageClients.txt_name.Value
                ws.Cells(nextLine, 2).Value = manageClients.txt_cnpj.Value
                ws.Cells(nextLine, 3).Value = manageClients.txt_street.Value
                ws.Cells(nextLine, 4).Value = manageClients.txt_number.Value
                ws.Cells(nextLine, 5).Value = manageClients.txt_nbhood.Value
                ws.Cells(nextLine, 6).Value = manageClients.txt_zipcode.Value
                ws.Cells(nextLine, 7).Value = manageClients.txt_city.Value
                ws.Cells(nextLine, 8).Value = manageClients.comb_state.Value
                ws.Cells(nextLine, 9).Value = manageClients.txt_phone_number.Value
                ws.Cells(nextLine, 10).Value = manageClients.txt_buyer.Value
                ws.Cells(nextLine, 11).Value = manageClients.txt_email.Value
    
                ' Desabilitar os campos após a adição
                manageClients.txt_name.Enabled = False
                manageClients.txt_cnpj.Enabled = False
                manageClients.txt_street.Enabled = False
                manageClients.txt_number.Enabled = False
                manageClients.txt_nbhood.Enabled = False
                manageClients.txt_zipcode.Enabled = False
                manageClients.txt_city.Enabled = False
                manageClients.comb_state.Enabled = False
                manageClients.txt_phone_number.Enabled = False
                manageClients.txt_buyer.Enabled = False
                manageClients.txt_email.Enabled = False


    
                ' Atualizar a lista de usuários
                Call def_load_list_clients
                
                manageClients.btn_add.BackColor = RGB(25, 86, 180)
                manageClients.btn_home.BackColor = RGB(25, 86, 180)
                manageClients.btn_home.Caption = "HOME"
                
            End If
            
        End If
        
    End If

End If

End Sub
