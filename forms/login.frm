VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} login 
   Caption         =   "DEAL FORGE"
   ClientHeight    =   9506
   ClientLeft      =   91
   ClientTop       =   406
   ClientWidth     =   9821
   OleObjectBlob   =   "login.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub UserForm_Initialize()

Call def_style_login

login.txt_username.Value = "admin"
login.txt_password.Value = "admin"

End Sub

Private Sub btn_exit_Click()

Call def_really_exit

End Sub

Private Sub btn_login_Click()

Call def_auth_user

End Sub


