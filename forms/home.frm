VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} home 
   Caption         =   "DEAL FORGE"
   ClientHeight    =   9506
   ClientLeft      =   91
   ClientTop       =   406
   ClientWidth     =   14819
   OleObjectBlob   =   "home.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "home"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btn_exit_Click()

Call def_really_exit

End Sub

Private Sub btn_logout_Click()

Call def_logout

End Sub

Private Sub btn_manage_clients_Click()

home.Hide
manageClients.Show

End Sub

Private Sub btn_manage_products_Click()

home.Hide
manageProducts.Show

End Sub

Private Sub btn_manage_users_Click()

Call def_is_admin

End Sub

Private Sub btn_new_deal_Click()

home.Hide
newDeal.Show

End Sub

Private Sub UserForm_Initialize()

Call def_style_home

End Sub
