VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} manageClients 
   Caption         =   "DEAL FORGE"
   ClientHeight    =   9814
   ClientLeft      =   91
   ClientTop       =   406
   ClientWidth     =   14987
   OleObjectBlob   =   "manageClients.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "manageClients"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btn_clean_Click()

If manageClients.btn_home.BackColor = RGB(255, 0, 0) Then

    MsgBox "Finalize a operação atual!", vbCritical, "DEAL FORGE"

Else

    Call def_load_list_clients
    manageClients.txt_search.Value = "" _

End If

End Sub

Private Sub btn_delete_Click()

Dim really As VbMsgBoxResult

If manageClients.txt_name.Value = "" _
Or manageClients.txt_name.Value = "NOME" _
Then

    MsgBox "Nenhum produto selecionado!", vbExclamation, "DEAL FORGE"

ElseIf manageClients.btn_home.BackColor = RGB(255, 0, 0) Then

    MsgBox "Finalize a operação atual antes de deletar um produto!", vbCritical, "DEAL FORGE"

Else

    really = MsgBox("Tem certeza que deseja DELETAR o produto selecionado?", vbYesNo + vbCritical, "DEAL FORGE")
    
    If really = vbYes Then
    
        Call def_delete_client
    
    End If

End If

End Sub

Private Sub btn_modify_Click()

If manageClients.btn_add.BackColor = RGB(0, 176, 80) Then

    MsgBox "Saia do modo Adicionar Cliente antes de executar esta tarefa!", vbCritical, "DEAL FORGE"
    
    Exit Sub

ElseIf manageClients.txt_name = "" Or manageClients.txt_name = "NOME" Then

    MsgBox "Nenhum cliente selecionado!", vbExclamation, "DEAL FORGE"

Else

    manageClients.btn_home.BackColor = RGB(255, 0, 0)

    Call def_modify_client
        
End If

End Sub

Private Sub btn_search_Click()

If manageClients.btn_home.BackColor = RGB(255, 0, 0) Then

    MsgBox "Finalize a operação atual!", vbCritical, "DEAL FORGE"

Else

    Call def_search_client
    
End If

End Sub

Private Sub list_clients_Change()

Call def_change_listbox_manageClients

End Sub

Private Sub btn_add_Click()
    
Call def_add_client

End Sub

Private Sub btn_home_Click()

If manageClients.btn_home.BackColor = RGB(255, 0, 0) Then

    Unload manageClients
    manageClients.Show

Else

    manageClients.Hide
    home.Show

End If



End Sub

Private Sub UserForm_Initialize()

Call def_style_manageClients
Call def_load_list_clients
Call def_load_comb_state

End Sub
