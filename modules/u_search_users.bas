Attribute VB_Name = "u_search_users"
Sub def_search_users()

Dim ws As Worksheet
Dim lastRow As Long
Dim i As Long
Dim matchCount As Long
Dim listData() As Variant
Dim listBox As MSForms.listBox
Dim wantedValue As String

' Definindo a planilha e o valor de pesquisa
Set ws = ThisWorkbook.Sheets("users")
wantedValue = manageUsers.txt_search.Value

' Encontrando a �ltima linha da planilha com dados
lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).row

' Definindo a ListBox
Set listBox = manageUsers.list_users

' Limpando a ListBox antes de preencher
listBox.Clear

' Inicializando o contador de correspond�ncias
matchCount = 0

' Adicionando o cabe�alho na ListBox
listBox.AddItem "NOME" ' Linha 1 - Cabe�alho
listBox.List(0, 1) = "NOME DE USU�RIO"
listBox.List(0, 2) = "SENHA"
listBox.List(0, 3) = "CLASSE"

' Primeiro loop para contar as correspond�ncias, come�ando a partir da linha 2 (ignorando o cabe�alho)
For i = 2 To lastRow ' Come�ando da linha 2

    ' Usando InStr para buscar partes do texto
    If InStr(1, ws.Cells(i, "A").Value, wantedValue, vbTextCompare) > 0 Or _
       InStr(1, ws.Cells(i, "B").Value, wantedValue, vbTextCompare) > 0 Or _
       InStr(1, ws.Cells(i, "D").Value, wantedValue, vbTextCompare) > 0 Then
       
        matchCount = matchCount + 1
        
        ' Adicionando os dados encontrados � ListBox a partir da segunda linha
        listBox.AddItem ws.Cells(i, "A").Value
        listBox.List(matchCount, 1) = ws.Cells(i, "B").Value
        listBox.List(matchCount, 2) = ws.Cells(i, "C").Value
        listBox.List(matchCount, 3) = ws.Cells(i, "D").Value
        
    End If
    
Next i

End Sub

