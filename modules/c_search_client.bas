Attribute VB_Name = "c_search_client"
Sub def_search_client()

Dim ws As Worksheet
Dim lastRow As Long
Dim i As Long
Dim matchCount As Long
Dim listData() As Variant
Dim listBox As MSForms.listBox
Dim wantedValue As String

' Definindo a planilha e o valor de pesquisa
Set ws = ThisWorkbook.Sheets("clients")
wantedValue = manageClients.txt_search.Value

' Encontrando a última linha da planilha com dados
lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).row

' Definindo a ListBox
Set listBox = manageClients.list_clients

' Limpando a ListBox antes de preencher
listBox.Clear

' Inicializando o contador de correspondências
matchCount = 0

' Adicionando o cabeçalho na ListBox

listBox.AddItem "NOMA DA EMPRESA" ' Linha 1 - Cabeçalho
listBox.List(0, 1) = "CNPJ"
listBox.List(0, 2) = "RUA"
listBox.List(0, 3) = "NÚMERO"
listBox.List(0, 4) = "BAIRRO"
listBox.List(0, 5) = "CEP"
listBox.List(0, 6) = "CIDADE"
listBox.List(0, 7) = "UF"
listBox.List(0, 8) = "TELEFONE"
listBox.List(0, 9) = "COMPRADOR"
listBox.List(0, 10) = "EMAIL"

' Primeiro loop para contar as correspondências, começando a partir da linha 2 (ignorando o cabeçalho)
For i = 2 To lastRow ' Começando da linha 2

    ' Usando InStr para buscar partes do texto
    If InStr(1, ws.Cells(i, "A").Value, wantedValue, vbTextCompare) > 0 Or _
       InStr(1, ws.Cells(i, "B").Value, wantedValue, vbTextCompare) > 0 Or _
       InStr(1, ws.Cells(i, "C").Value, wantedValue, vbTextCompare) > 0 Or _
       InStr(1, ws.Cells(i, "D").Value, wantedValue, vbTextCompare) > 0 Or _
       InStr(1, ws.Cells(i, "E").Value, wantedValue, vbTextCompare) > 0 Or _
       InStr(1, ws.Cells(i, "F").Value, wantedValue, vbTextCompare) > 0 Or _
       InStr(1, ws.Cells(i, "G").Value, wantedValue, vbTextCompare) > 0 Or _
       InStr(1, ws.Cells(i, "H").Value, wantedValue, vbTextCompare) > 0 Or _
       InStr(1, ws.Cells(i, "I").Value, wantedValue, vbTextCompare) > 0 Or _
       InStr(1, ws.Cells(i, "J").Value, wantedValue, vbTextCompare) > 0 Or _
       InStr(1, ws.Cells(i, "K").Value, wantedValue, vbTextCompare) > 0 Then
       
        matchCount = matchCount + 1
        
        ' Adicionando os dados encontrados à ListBox a partir da segunda linha
        listBox.AddItem ws.Cells(i, "A").Value
        listBox.List(matchCount, 1) = ws.Cells(i, "B").Value
        listBox.List(matchCount, 2) = ws.Cells(i, "C").Value
        listBox.List(matchCount, 3) = ws.Cells(i, "D").Value
        listBox.List(matchCount, 4) = ws.Cells(i, "E").Value
        listBox.List(matchCount, 5) = ws.Cells(i, "F").Value
        listBox.List(matchCount, 6) = ws.Cells(i, "G").Value
        listBox.List(matchCount, 7) = ws.Cells(i, "H").Value
        listBox.List(matchCount, 8) = ws.Cells(i, "I").Value
        listBox.List(matchCount, 9) = ws.Cells(i, "J").Value
        listBox.List(matchCount, 10) = ws.Cells(i, "K").Value
        
    End If
    
Next i

End Sub




