VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} confirmPassword 
   Caption         =   "DEAL FORGE"
   ClientHeight    =   2863
   ClientLeft      =   91
   ClientTop       =   406
   ClientWidth     =   4298
   OleObjectBlob   =   "confirmPassword.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "confirmPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btn_cancel_Click()

Unload confirmPassword

End Sub

Private Sub btn_confirm_Click()

Call def_confirm_password

End Sub

Private Sub UserForm_Initialize()

Call def_style_confirmPassword

confirmPassword.txt_password.Value = "admin"

End Sub
