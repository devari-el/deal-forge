Attribute VB_Name = "c_change_listbox_manageClients"
Sub def_change_listbox_manageClients()

    Dim listBox As MSForms.listBox
    
    ' Associa os controles
    With manageClients
        Set listBox = .list_clients
        
        ' Se o campo nome estiver habilitado, sair
        If .txt_name.Enabled Then Exit Sub
        
        ' Verifica se há um item selecionado (excluindo o cabeçalho)
        If listBox.ListIndex > 0 Then
            ' Preenche os campos com base na seleção
            .txt_name.Value = listBox.List(listBox.ListIndex, 0)
            .txt_cnpj.Value = listBox.List(listBox.ListIndex, 1)
            .txt_street.Value = listBox.List(listBox.ListIndex, 2)
            .txt_number.Value = listBox.List(listBox.ListIndex, 3)
            .txt_nbhood.Value = listBox.List(listBox.ListIndex, 4)
            .txt_zipcode.Value = listBox.List(listBox.ListIndex, 5)
            .txt_city.Value = listBox.List(listBox.ListIndex, 6)
            .comb_state.Value = listBox.List(listBox.ListIndex, 7)
            .txt_phone_number.Value = listBox.List(listBox.ListIndex, 8)
            .txt_buyer.Value = listBox.List(listBox.ListIndex, 9)
            .txt_email.Value = listBox.List(listBox.ListIndex, 10)
        Else
            ' Limpa os campos se nenhum item válido estiver selecionado
            .txt_name.Value = ""
            .txt_cnpj.Value = ""
            .txt_street.Value = ""
            .txt_number.Value = ""
            .txt_nbhood.Value = ""
            .txt_zipcode.Value = ""
            .txt_city.Value = ""
            .comb_state.Value = ""
            .txt_phone_number.Value = ""
            .txt_buyer.Value = ""
            .txt_email.Value = ""
        End If
    End With
    
End Sub


