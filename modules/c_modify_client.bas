Attribute VB_Name = "c_modify_client"
Sub def_modify_client()

Dim ws As Worksheet
Dim selectedIndex As Long
Dim rowToEdit As Long
Dim existingRow As Long
Dim i As Long
Dim cnpj As String
Dim cnpjAlreadyExists As Boolean

Set ws = ThisWorkbook.Sheets("Clients")

' Verificar se há campos vazios
If manageClients.txt_name.Value = "" _
Or manageClients.txt_cnpj.Value = "" _
Or manageClients.txt_street.Value = "" _
Or manageClients.txt_number.Value = "" _
Or manageClients.txt_nbhood.Value = "" _
Or manageClients.txt_zipcode.Value = "" _
Or manageClients.txt_city.Value = "" _
Or manageClients.comb_state.Value = "" _
Or manageClients.txt_phone_number.Value = "" _
Or manageClients.txt_buyer.Value = "" _
Or manageClients.txt_email.Value = "" Then
   
    MsgBox "Existem campos vazios!", vbInformation, "DEAL FORGE"
    
    Exit Sub
    
End If

' Ativar modo de edição
If manageClients.txt_name.Enabled = False Then

    ' Habilitar os campos de entrada
    manageClients.txt_name.Enabled = True
    manageClients.txt_cnpj.Enabled = True
    manageClients.txt_street.Enabled = True
    manageClients.txt_number.Enabled = True
    manageClients.txt_nbhood.Enabled = True
    manageClients.txt_zipcode.Enabled = True
    manageClients.txt_city.Enabled = True
    manageClients.comb_state.Enabled = True
    manageClients.txt_phone_number.Enabled = True
    manageClients.txt_buyer.Enabled = True
    manageClients.txt_email.Enabled = True

    manageClients.btn_modify.BackColor = RGB(0, 176, 80)
    manageClients.btn_home.BackColor = RGB(255, 0, 0)
    manageClients.btn_home.Caption = "CANCEL"
    
Else
    ' Obter índice do item selecionado
    selectedIndex = manageClients.list_clients.ListIndex

    ' Obter a linha correspondente na planilha
    rowToEdit = selectedIndex + 1

    ' Verificar se o código já existe na coluna B
    cnpj = manageClients.txt_cnpj.Value
    cnpjAlreadyExists = False

    For i = 2 To ws.Cells(ws.Rows.count, "B").End(xlUp).row
        If ws.Cells(i, 2).Value = cnpj And i <> rowToEdit Then
            cnpjAlreadyExists = True
            Exit For
        End If
    Next i

    If cnpjAlreadyExists Then
        MsgBox "Este CNPJ já está registrado!", vbCritical, "DEAL FORGE"
        Exit Sub
    End If

    ' Atualizar os dados na planilha
    ws.Cells(rowToEdit, 1).Value = manageClients.txt_name.Value
    ws.Cells(rowToEdit, 2).Value = manageClients.txt_cnpj.Value
    ws.Cells(rowToEdit, 3).Value = manageClients.txt_street.Value
    ws.Cells(rowToEdit, 4).Value = manageClients.txt_number.Value
    ws.Cells(rowToEdit, 5).Value = manageClients.txt_nbhood.Value
    ws.Cells(rowToEdit, 6).Value = manageClients.txt_zipcode.Value
    ws.Cells(rowToEdit, 7).Value = manageClients.txt_city.Value
    ws.Cells(rowToEdit, 8).Value = manageClients.comb_state.Value
    ws.Cells(rowToEdit, 9).Value = manageClients.txt_phone_number.Value
    ws.Cells(rowToEdit, 10).Value = manageClients.txt_buyer.Value
    ws.Cells(rowToEdit, 11).Value = manageClients.txt_email.Value


    ' Atualizar a ListBox
    manageClients.list_clients.List(selectedIndex, 0) = manageClients.txt_name.Value
    manageClients.list_clients.List(selectedIndex, 1) = manageClients.txt_cnpj.Value
    manageClients.list_clients.List(selectedIndex, 2) = manageClients.txt_street.Value
    manageClients.list_clients.List(selectedIndex, 3) = manageClients.txt_number.Value
    manageClients.list_clients.List(selectedIndex, 4) = manageClients.txt_nbhood.Value
    manageClients.list_clients.List(selectedIndex, 5) = manageClients.txt_zipcode.Value
    manageClients.list_clients.List(selectedIndex, 6) = manageClients.txt_city.Value
    manageClients.list_clients.List(selectedIndex, 7) = manageClients.comb_state.Value
    manageClients.list_clients.List(selectedIndex, 8) = manageClients.txt_phone_number.Value
    manageClients.list_clients.List(selectedIndex, 9) = manageClients.txt_buyer.Value
    manageClients.list_clients.List(selectedIndex, 10) = manageClients.txt_email.Value

    ' Desabilitar os campos após a adição
    manageClients.txt_name.Enabled = False
    manageClients.txt_cnpj.Enabled = False
    manageClients.txt_street.Enabled = False
    manageClients.txt_number.Enabled = False
    manageClients.txt_nbhood.Enabled = False
    manageClients.txt_zipcode.Enabled = False
    manageClients.txt_city.Enabled = False
    manageClients.comb_state.Enabled = False
    manageClients.txt_phone_number.Enabled = False
    manageClients.txt_buyer.Enabled = False
    manageClients.txt_email.Enabled = False


    manageClients.btn_modify.BackColor = RGB(25, 86, 180)
    manageClients.btn_home.BackColor = RGB(25, 86, 180)
    manageClients.btn_home.Caption = "HOME"
    
    Call def_load_list_clients
        
End If



End Sub




