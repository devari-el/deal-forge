VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} manageDeals 
   Caption         =   "DEAL FORGE"
   ClientHeight    =   7005
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12330
   OleObjectBlob   =   "manageDeals.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "manageDeals"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btn_home_Click()

Unload manageDeals
home.Show

End Sub

Private Sub btn_update_Click()

Call def_update_status

End Sub

Private Sub list_deals_Change()

Call def_list_deals_change

End Sub

Private Sub UserForm_Initialize()

Call def_style_manageDeals
Call def_load_list_deals
Call def_load_comb_status

End Sub
