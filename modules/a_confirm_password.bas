Attribute VB_Name = "a_confirm_password"
Sub def_confirm_password()

Dim ws As Worksheet
Dim password As String
Dim isValid As Boolean

Set ws = ThisWorkbook.Sheets("users")

    password = confirmPassword.txt_password
    
    If password = ws.Cells(2, 8) Then
    
        isValid = True
        Unload confirmPassword
        
'---------------------------------------------------------------------
                
        If caller = "home" Then
            
            home.Hide
            manageUsers.Show
            
'---------------------------------------------------------------------
            
        ElseIf caller = "show_password" Then
        
            manageUsers.txt_password.PasswordChar = ""
          
'---------------------------------------------------------------------
          
        ElseIf caller = "add_user" Then
        
            Call def_add_user
        
'---------------------------------------------------------------------

        ElseIf caller = "modify_user" Then
        
            Call def_modify_user
            
'---------------------------------------------------------------------
        
         ElseIf caller = "delete_user" Then
        
            Call def_delete_user
            
'---------------------------------------------------------------------


        End If
        
    Else
        
        isValid = False
        MsgBox "Senha incorreta!", vbCritical, "DEAL FORGE"
        
    End If

caller = ""

End Sub
