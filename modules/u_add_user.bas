Attribute VB_Name = "u_add_user"
Sub def_add_user()

Dim name As String
Dim username As String
Dim password As String
Dim class As String
Dim ws As Worksheet
Dim lastRow As Long
Dim nextLine As Long
Dim optConfirm As Boolean
Dim cell As range
Dim userExists As Boolean

If manageUsers.btn_modify.BackColor = RGB(0, 176, 80) Then
    
    MsgBox "Saia do modo Alterar Usuário antes de executar esta tarefa!", vbCritical, "DEAL FORGE"

Else

    If manageUsers.txt_name.Enabled = False Then
    
        ' Habilitar os campos de entrada
        manageUsers.txt_name.Enabled = True
        manageUsers.txt_username.Enabled = True
        manageUsers.txt_password.Enabled = True
        manageUsers.opt_admin.Enabled = True
        manageUsers.opt_user.Enabled = True
    
        ' Limpar os campos de entrada
        manageUsers.txt_name = ""
        manageUsers.txt_username = ""
        manageUsers.txt_password = ""
        manageUsers.opt_admin = False
        manageUsers.opt_user = False
        
        manageUsers.btn_add.BackColor = RGB(0, 176, 80)
        manageUsers.btn_home.BackColor = RGB(255, 0, 0)
        manageUsers.btn_home.Caption = "CANCEL"
    
    Else
    
        Set ws = ThisWorkbook.Sheets("users")
        name = manageUsers.txt_name
        username = manageUsers.txt_username
        password = manageUsers.txt_password
    
        ' Determinar o tipo de usuário
        If manageUsers.opt_admin = True Then
            class = "admin"
            optConfirm = True
        ElseIf manageUsers.opt_user = True Then
            class = "user"
            optConfirm = True
        Else
            optConfirm = False
        End If
    
        ' Verificar campos obrigatórios
        If name = "" Or username = "" Or password = "" Or optConfirm = False Then
            MsgBox "Preencha todos os campos!", vbCritical, "DEAL FORGE"
        Else
            ' Verificar se o nome de usuário já existe
            userExists = False
            lastRow = ws.Cells(ws.Rows.count, "B").End(xlUp).row
            For Each cell In ws.range("B2:B" & lastRow) ' Verifica na coluna de usernames
                If cell.Value = username Then
                    userExists = True
                    Exit For
                End If
            Next cell
            
            If userExists Then
                MsgBox "O nome de usuário '" & username & "' já existe!", vbExclamation, "DEAL FORGE"
            Else
                ' Adicionar o novo usuário
                nextLine = lastRow + 1
                ws.Cells(nextLine, 1).Value = name
                ws.Cells(nextLine, 2).Value = username
                ws.Cells(nextLine, 3).Value = password
                ws.Cells(nextLine, 4).Value = class
    
                ' Desabilitar os campos após a adição
                manageUsers.txt_name.Enabled = False
                manageUsers.txt_username.Enabled = False
                manageUsers.txt_password.Enabled = False
                manageUsers.opt_admin.Enabled = False
                manageUsers.opt_user.Enabled = False
    
                ' Atualizar a lista de usuários
                Call def_load_list_users
                
                manageUsers.btn_add.BackColor = RGB(25, 86, 180)
                manageUsers.btn_home.BackColor = RGB(25, 86, 180)
                manageUsers.btn_home.Caption = "HOME"
                
            End If
            
        End If
        
    End If

End If

End Sub
