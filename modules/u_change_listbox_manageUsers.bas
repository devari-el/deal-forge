Attribute VB_Name = "u_change_listbox_manageUsers"
Sub def_change_listbox_manageUsers()

    Dim listBox As MSForms.listBox
    Dim tbName As MSForms.TextBox
    Dim tbUsername As MSForms.TextBox
    Dim tbPassword As MSForms.TextBox
    Dim optAdmin As MSForms.OptionButton
    Dim optUser As MSForms.OptionButton
    Dim typeUser As String
    
    ' Associar os controles do formulário
    With manageUsers
        Set listBox = .list_users
        Set tbName = .txt_name
        Set tbUsername = .txt_username
        Set tbPassword = .txt_password
        Set optAdmin = .opt_admin
        Set optUser = .opt_user
        
        ' Se o campo nome estiver habilitado, sair
        If tbName.Enabled Then Exit Sub
        
        ' Verifica se há um item selecionado (excluindo o cabeçalho)
        If listBox.ListIndex > 0 Then  ' Apenas índices maiores que 0 (cabeçalho é 0)
            ' Preencher os campos com base na seleção
            tbName.Value = listBox.List(listBox.ListIndex, 0)
            tbUsername.Value = listBox.List(listBox.ListIndex, 1)
            tbPassword.Value = listBox.List(listBox.ListIndex, 2)
            typeUser = listBox.List(listBox.ListIndex, 3)
            
            ' Definir o tipo de usuário
            optAdmin.Value = (typeUser = "admin")
            optUser.Value = (typeUser = "user")
        Else
            ' Limpar os campos se nenhum item válido estiver selecionado
            tbName.Value = ""
            tbUsername.Value = ""
            tbPassword.Value = ""
            optAdmin.Value = False
            optUser.Value = False
        End If
    End With
    
End Sub

