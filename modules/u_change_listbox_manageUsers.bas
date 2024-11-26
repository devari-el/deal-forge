Attribute VB_Name = "u_change_listbox_manageUsers"
Sub def_change_listbox_manageUsers()

    Dim listBox As MSForms.listBox
    Dim tbName As MSForms.TextBox
    Dim tbUsername As MSForms.TextBox
    Dim tbPassword As MSForms.TextBox
    Dim optAdmin As MSForms.OptionButton
    Dim optUser As MSForms.OptionButton
    Dim typeUser As String
    
    ' Associar os controles do formul�rio
    With manageUsers
        Set listBox = .list_users
        Set tbName = .txt_name
        Set tbUsername = .txt_username
        Set tbPassword = .txt_password
        Set optAdmin = .opt_admin
        Set optUser = .opt_user
        
        ' Se o campo nome estiver habilitado, sair
        If tbName.Enabled Then Exit Sub
        
        ' Verifica se h� um item selecionado (excluindo o cabe�alho)
        If listBox.ListIndex > 0 Then  ' Apenas �ndices maiores que 0 (cabe�alho � 0)
            ' Preencher os campos com base na sele��o
            tbName.Value = listBox.List(listBox.ListIndex, 0)
            tbUsername.Value = listBox.List(listBox.ListIndex, 1)
            tbPassword.Value = listBox.List(listBox.ListIndex, 2)
            typeUser = listBox.List(listBox.ListIndex, 3)
            
            ' Definir o tipo de usu�rio
            optAdmin.Value = (typeUser = "admin")
            optUser.Value = (typeUser = "user")
        Else
            ' Limpar os campos se nenhum item v�lido estiver selecionado
            tbName.Value = ""
            tbUsername.Value = ""
            tbPassword.Value = ""
            optAdmin.Value = False
            optUser.Value = False
        End If
    End With
    
End Sub

