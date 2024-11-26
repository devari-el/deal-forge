Attribute VB_Name = "l_auth_user"
Sub def_auth_user() 'Essa rotina valida login e senha

    Dim ws As Worksheet
    Dim username As String
    Dim password As String
    Dim lastRow As Long
    Dim i As Long
    Dim ValidUser As Boolean

    ' Definir a planilha onde estão os usuários e senhas
    Set ws = ThisWorkbook.Sheets("users")
    
    ' Capturar os valores inseridos pelo usuário
    username = login.txt_username.Value
    password = login.txt_password.Value
    
    ' Verificar se o campo de usuário ou senha está vazio
    If username = "" Or password = "" Then
        MsgBox "Por favor, preencha ambos os campos!", vbExclamation, "DEAL FORGE"
        Exit Sub
    End If
    
    ' Encontrar a última linha da coluna B (Usuários)
    lastRow = ws.Cells(ws.Rows.count, "B").End(xlUp).row
    
    ' Inicializar a variável que vai verificar se o usuário é válido
    ValidUser = False
    
    ' Loop para percorrer as linhas e verificar se o usuário e senha são válidos
    For i = 2 To lastRow ' Começar da linha 2, assumindo que a linha 1 tem cabeçalhos
        If ws.Cells(i, 2).Value = username And ws.Cells(i, 3).Value = password Then
            ValidUser = True
            ws.Cells(2, 6) = ws.Cells(i, 1).Value
            ws.Cells(2, 7) = ws.Cells(i, 4).Value
            ws.Cells(2, 8) = ws.Cells(i, 3).Value
            Exit For
        End If
    Next i
    
    ' Se o usuário for encontrado e a senha for correta
    If ValidUser Then
    
        Unload login
        home.Show

    Else
        MsgBox "Usuário ou senha inválidos. Tente novamente.", vbCritical, "DEAL FORGE"
    End If

End Sub
