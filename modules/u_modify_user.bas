Attribute VB_Name = "u_modify_user"
Sub def_modify_user()

Dim ws As Worksheet
Dim selectedIndex As Long
Dim rowToEdit As Long
Dim class As String
Dim existingRow As Long
Dim i As Long
Dim username As String
Dim usernameAlreadyExists As Boolean

Set ws = ThisWorkbook.Sheets("users")

' Verificar se há campos vazios
If manageUsers.txt_name = "" Or _
   manageUsers.txt_username = "" Or _
   manageUsers.txt_password = "" Or _
   (manageUsers.opt_admin = False And manageUsers.opt_user = False) Then
   
    MsgBox "Existem campos vazios!", vbInformation, "DEAL FORGE"
    
    Exit Sub
    
End If

' Ativar modo de edição
If manageUsers.txt_name.Enabled = False Then

    manageUsers.txt_name.Enabled = True
    manageUsers.txt_username.Enabled = True
    manageUsers.txt_password.Enabled = True
    manageUsers.opt_admin.Enabled = True
    manageUsers.opt_user.Enabled = True

    manageUsers.btn_modify.BackColor = RGB(0, 176, 80)
    manageUsers.btn_home.BackColor = RGB(255, 0, 0)
    manageUsers.btn_home.Caption = "CANCEL"
    
Else
    ' Obter índice do item selecionado
    selectedIndex = manageUsers.list_users.ListIndex

    ' Obter a linha correspondente na planilha
    rowToEdit = selectedIndex + 1

    ' Verificar se o nome de usuário já existe na coluna B
    username = manageUsers.txt_username.Value
    usernameAlreadyExists = False

    For i = 2 To ws.Cells(ws.Rows.count, "B").End(xlUp).row
        If ws.Cells(i, 2).Value = username And i <> rowToEdit Then
            usernameAlreadyExists = True
            Exit For
        End If
    Next i

    If usernameAlreadyExists Then
        MsgBox "O nome de usuário já existe!", vbCritical, "DEAL FORGE"
        Exit Sub
    End If

    ' Determinar o tipo de usuário
    If manageUsers.opt_admin = True Then
        class = "admin"
    Else
        class = "user"
    End If

    ' Atualizar os dados na planilha
    ws.Cells(rowToEdit, 1).Value = manageUsers.txt_name.Value
    ws.Cells(rowToEdit, 2).Value = manageUsers.txt_username.Value
    ws.Cells(rowToEdit, 3).Value = manageUsers.txt_password.Value
    ws.Cells(rowToEdit, 4).Value = class

    ' Atualizar a ListBox
    manageUsers.list_users.List(selectedIndex, 0) = manageUsers.txt_name.Value
    manageUsers.list_users.List(selectedIndex, 1) = manageUsers.txt_username.Value
    manageUsers.list_users.List(selectedIndex, 2) = manageUsers.txt_password.Value
    manageUsers.list_users.List(selectedIndex, 3) = class

    ' Desabilitar os campos após edição
    manageUsers.txt_name.Enabled = False
    manageUsers.txt_username.Enabled = False
    manageUsers.txt_password.Enabled = False
    manageUsers.opt_admin.Enabled = False
    manageUsers.opt_user.Enabled = False

    manageUsers.btn_modify.BackColor = RGB(25, 86, 180)
    manageUsers.btn_home.BackColor = RGB(25, 86, 180)
    manageUsers.btn_home.Caption = "HOME"
    
    Call def_load_list_users
        
End If



End Sub
