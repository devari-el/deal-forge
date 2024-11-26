VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} manageProducts 
   Caption         =   "DEAL FORGE"
   ClientHeight    =   9639
   ClientLeft      =   91
   ClientTop       =   406
   ClientWidth     =   14868
   OleObjectBlob   =   "manageProducts.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "manageProducts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btn_clean_Click()

If manageProducts.btn_home.BackColor = RGB(255, 0, 0) Then

    MsgBox "Finalize a opera��o atual!", vbCritical, "DEAL FORGE"

Else

    Call def_load_list_products
    manageProducts.txt_search.Value = "" _

End If

End Sub

Private Sub btn_delete_Click()

Dim really As VbMsgBoxResult

If manageProducts.txt_name.Value = "" _
Or manageProducts.txt_name.Value = "NOME" _
Then

    MsgBox "Nenhum produto selecionado!", vbExclamation, "DEAL FORGE"

ElseIf manageProducts.btn_home.BackColor = RGB(255, 0, 0) Then

    MsgBox "Finalize a opera��o atual antes de deletar um produto!", vbCritical, "DEAL FORGE"

Else

    really = MsgBox("Tem certeza que deseja DELETAR o produto selecionado?", vbYesNo + vbCritical, "DEAL FORGE")
    
    If really = vbYes Then
    
        Call def_delete_product
    
    End If

End If

End Sub

Private Sub btn_search_Click()

If manageProducts.btn_home.BackColor = RGB(255, 0, 0) Then

    MsgBox "Finalize a opera��o atual!", vbCritical, "DEAL FORGE"

Else

    Call def_search_product
    
End If

End Sub

Private Sub list_products_Change()

Call def_change_listbox_manageProducts

End Sub

Private Sub btn_add_Click()
    
    Call def_add_product

End Sub

Private Sub btn_home_Click()

If manageProducts.btn_home.BackColor = RGB(255, 0, 0) Then

    Unload manageProducts
    manageProducts.Show

Else

    manageProducts.Hide
    home.Show

End If



End Sub

Private Sub btn_modify_Click()

If manageProducts.btn_add.BackColor = RGB(0, 176, 80) Then

    MsgBox "Saia do modo Adicionar Produto antes de executar esta tarefa!", vbCritical, "DEAL FORGE"
    
    Exit Sub

ElseIf manageProducts.txt_name = "" Or manageProducts.txt_name = "NOME" Then

    MsgBox "Nenhum produto selecionado!", vbExclamation, "DEAL FORGE"

Else

    manageProducts.btn_home.BackColor = RGB(255, 0, 0)

    Call def_modify_product
        
End If

End Sub

Private Sub UserForm_Initialize()

Call def_style_manageProducts
Call def_load_list_products

End Sub
