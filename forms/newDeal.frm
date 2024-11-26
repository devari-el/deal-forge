VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} newDeal 
   Caption         =   "DEAL FORGE"
   ClientHeight    =   8835
   ClientLeft      =   91
   ClientTop       =   406
   ClientWidth     =   15001
   OleObjectBlob   =   "newDeal.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "newDeal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btn_add_Click()

    Call def_add_product_deal

End Sub

Private Sub btn_clean_Click()

Call def_load_list_products_nd
newDeal.txt_search.Value = ""

End Sub

Private Sub btn_deal_Click()

If validateData() Then
    Call def_export_deal
    Call def_clean_layout
Else
    Call fillDataMsg
End If

End Sub

Private Sub btn_del_Click()

    Call def_delete_item

End Sub

Private Sub btn_home_Click()

Unload newDeal
home.Show

End Sub

Private Sub btn_search_Click()

    caller = "nd"

    Call def_search_product
    
End Sub

Private Sub label_obs_Click()

End Sub

Private Sub list_deal_Click()

End Sub

Private Sub list_products_Click()

End Sub

Private Sub txt_descount_Change()

End Sub

Private Sub UserForm_Initialize()

Call def_style_newDeal
Call def_load_data

End Sub
