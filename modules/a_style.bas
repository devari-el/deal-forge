Attribute VB_Name = "a_style"

Sub def_style_login() 'Esta rotina define o design da login page

Dim dir_logo As String

dir_logo = ThisWorkbook.path & "\style\logo.jpg"

login.logo.Picture = LoadPicture(dir_logo)
login.BackColor = RGB(25, 86, 180)
login.btn_login.BackColor = RGB(25, 86, 180)

End Sub

Sub def_style_home() 'Esta rotina define o design da home page

Dim dir_logo As String
Dim ws As Worksheet
Dim username As String

Set ws = ThisWorkbook.Sheets("users")

dir_logo = ThisWorkbook.path & "\style\logo.jpg"

username = ws.range("F2").Value

home.label_username.Caption = "Bem-vindo(a), " & username & "!"

home.logo.Picture = LoadPicture(dir_logo)
home.BackColor = RGB(25, 86, 180)
home.btn_new_deal.BackColor = RGB(25, 86, 180)
home.btn_manage_deal.BackColor = RGB(25, 86, 180)
home.btn_manage_clients.BackColor = RGB(25, 86, 180)
home.btn_manage_products.BackColor = RGB(25, 86, 180)
home.btn_manage_users.BackColor = RGB(25, 86, 180)
home.btn_logout.BackColor = RGB(25, 86, 180)

End Sub

Sub def_style_manageUsers() 'Esta rotina define o design da manageUsers page

manageUsers.BackColor = RGB(25, 86, 180)
manageUsers.btn_home.BackColor = RGB(25, 86, 180)
manageUsers.btn_add.BackColor = RGB(25, 86, 180)
manageUsers.btn_delete.BackColor = RGB(25, 86, 180)
manageUsers.btn_modify.BackColor = RGB(25, 86, 180)
manageUsers.btn_search.BackColor = RGB(25, 86, 180)
manageUsers.btn_clean.BackColor = RGB(25, 86, 180)
manageUsers.btn_show.BackColor = RGB(25, 86, 180)
manageUsers.frame_class.BackColor = RGB(25, 86, 180)
manageUsers.frame_users.BackColor = RGB(25, 86, 180)


End Sub

Sub def_style_manageProducts() 'Esta rotina define o design da manageProducts page

manageProducts.BackColor = RGB(25, 86, 180)
manageProducts.btn_home.BackColor = RGB(25, 86, 180)
manageProducts.btn_add.BackColor = RGB(25, 86, 180)
manageProducts.btn_delete.BackColor = RGB(25, 86, 180)
manageProducts.btn_modify.BackColor = RGB(25, 86, 180)
manageProducts.btn_search.BackColor = RGB(25, 86, 180)
manageProducts.btn_clean.BackColor = RGB(25, 86, 180)
manageProducts.frame_type.BackColor = RGB(25, 86, 180)
manageProducts.frame_products.BackColor = RGB(25, 86, 180)

End Sub

Sub def_style_manageClients() 'Esta rotina define o design da manageProducts page

manageClients.BackColor = RGB(25, 86, 180)
manageClients.btn_home.BackColor = RGB(25, 86, 180)
manageClients.btn_add.BackColor = RGB(25, 86, 180)
manageClients.btn_delete.BackColor = RGB(25, 86, 180)
manageClients.btn_modify.BackColor = RGB(25, 86, 180)
manageClients.btn_search.BackColor = RGB(25, 86, 180)
manageClients.btn_clean.BackColor = RGB(25, 86, 180)
manageClients.frame_adress.BackColor = RGB(25, 86, 180)
manageClients.frame_clients.BackColor = RGB(25, 86, 180)

End Sub

Sub def_style_confirmPassword() 'Esta rotina define o design da confirmPassword page

confirmPassword.BackColor = RGB(25, 86, 180)
confirmPassword.btn_confirm.BackColor = RGB(25, 86, 180)

End Sub

Sub def_style_newDeal() 'Esta rotina define o design da newDeal page

newDeal.BackColor = RGB(25, 86, 180)
newDeal.btn_home.BackColor = RGB(25, 86, 180)
newDeal.btn_add.BackColor = RGB(25, 86, 180)
newDeal.btn_clean.BackColor = RGB(25, 86, 180)
newDeal.btn_deal.BackColor = RGB(25, 86, 180)
newDeal.btn_del.BackColor = RGB(255, 0, 0)
newDeal.btn_search.BackColor = RGB(25, 86, 180)
newDeal.frame_deal.BackColor = RGB(25, 86, 180)

End Sub
