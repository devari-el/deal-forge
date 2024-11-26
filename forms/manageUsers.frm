VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} manageUsers 
   Caption         =   "DEAL FORGE"
   ClientHeight    =   9506
   ClientLeft      =   91
   ClientTop       =   406
   ClientWidth     =   9849
   OleObjectBlob   =   "manageUsers.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "manageUsers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btn_add_Click()

If manageUsers.btn_home.BackColor = RGB(255, 0, 0) Then
    
    Call def_add_user

Else

    caller = "add_user"
    confirmPassword.Show

End If
    
End Sub

Private Sub btn_clean_Click()

If manageUsers.btn_home.BackColor = RGB(255, 0, 0) Then

    MsgBox "Finalize a operação atual!", vbCritical, "DEAL FORGE"

Else

    Call def_load_list_users
    manageUsers.txt_search.Value = "" _

End If

End Sub

Private Sub btn_delete_Click()

Dim really As VbMsgBoxResult

If manageUsers.txt_name.Value = "" _
Or manageUsers.txt_name.Value = "NOME" _
Then

    MsgBox "Nenhum usuário selecionado!", vbExclamation, "DEAL FORGE"

ElseIf manageUsers.btn_home.BackColor = RGB(255, 0, 0) Then

    MsgBox "Finalize a operação atual antes de deletar um usuário!", vbCritical, "DEAL FORGE"

Else

    really = MsgBox("Tem certeza que deseja DELETAR o usuário selecionado?", vbYesNo + vbCritical, "DEAL FORGE")
    
    If really = vbYes Then
    
        caller = "delete_user"
        confirmPassword.Show
    
    End If

End If

End Sub

Private Sub btn_modify_Click()

If manageUsers.btn_add.BackColor = RGB(0, 176, 80) Then

    MsgBox "Saia do modo Adicionar Usuário antes de executar esta tarefa!", vbCritical, "DEAL FORGE"
    
    Exit Sub

ElseIf manageUsers.txt_name = "" Or manageUsers.txt_name = "NOME" Then

    MsgBox "Nenhum usuário selecionado!", vbExclamation, "DEAL FORGE"

Else

    If manageUsers.btn_home.BackColor = RGB(255, 0, 0) Then
    
        Call def_modify_user
        
    Else
    
        caller = "modify_user"
        confirmPassword.Show
    
    End If
    
End If

End Sub

Private Sub btn_search_Click()

If manageUsers.btn_home.BackColor = RGB(255, 0, 0) Then

    MsgBox "Finalize a operação atual!", vbCritical, "DEAL FORGE"

Else

    Call def_search_users
    
End If

End Sub

Private Sub UserForm_Initialize()

Call def_style_manageUsers
Call def_load_list_users
    
End Sub

Private Sub btn_home_Click()

If manageUsers.btn_home.BackColor = RGB(255, 0, 0) Then

    Unload manageUsers
    manageUsers.Show

Else

    manageUsers.Hide
    home.Show

End If

End Sub

Private Sub list_users_Change()

Call def_change_listbox_manageUsers

End Sub

Private Sub btn_show_Click()

If manageUsers.txt_password.PasswordChar = "•" Then
    
    caller = "show_password"
    confirmPassword.Show

Else

    manageUsers.txt_password.PasswordChar = "•"

End If

End Sub
