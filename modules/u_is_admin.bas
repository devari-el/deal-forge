Attribute VB_Name = "u_is_admin"
Sub def_is_admin()

Dim is_admin As Boolean
Dim ws As Worksheet

Set ws = ThisWorkbook.Sheets("Users")

If ws.Cells(2, 7).Value = "admin" Then

    caller = "home"
    confirmPassword.Show

Else: MsgBox "Somente a Classe ADMIN pode abrir esta sessão!", vbCritical, "DEAL FORGE"

End If

End Sub
