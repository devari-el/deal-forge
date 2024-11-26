Attribute VB_Name = "nd_validate_data"
Public Function fillDataMsg() As VbMsgBoxResult

fillData = MsgBox("Preencha todos os campos obrigatórios", vbExclamation, "DEAL FORGE")
    
End Function
Public Function validateData() As Boolean

Dim cmbClient As String
Dim cmbDelivery As String
Dim cmbTerm As String
Dim cmbConditions As String

cmbClient = newDeal.comb_client.Value
cmbDelivery = newDeal.comb_delivery.Value
cmbTerm = newDeal.comb_term.Value
cmbConditions = newDeal.comb_conditions.Value

'Definindo o array de campos do formulário
Dim dataArray(1 To 4) As String
dataArray(1) = cmbClient
dataArray(2) = cmbDelivery
dataArray(3) = cmbTerm
dataArray(4) = cmbConditions

'Checando se algum campo obrigatório está vazio
Dim isDataValid As Boolean
isDataValid = True

Dim i As Integer
For i = LBound(dataArray) To UBound(dataArray)
    If dataArray(i) = "" Then
        isDataValid = False
    End If
Next i

validateData = isDataValid

End Function

