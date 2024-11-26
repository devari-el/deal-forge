Attribute VB_Name = "nd_export_deal"
'Função para obter a pasta de orcamentos
Function getDealsFolder() As String

Dim currentPath As String
Dim dealsFolder As String

' Obtém o caminho completo do arquivo atual
currentPath = ThisWorkbook.path

' Cria o caminho para a pasta de orçamentos
dealsFolder = currentPath & "\Deals\"

' Cria a pasta de orcamentos se não existir
If Dir(dealsFolder, vbDirectory) = "" Then
    MkDir dealsFolder
End If

getDealsFolder = dealsFolder

End Function

'Função para obter o número do orcamento + ano atual
Function getDealNumber(folderPath As String) As String

Dim file As Variant
Dim i As Integer

If Right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"
file = Dir(folderPath)

While (file <> "")
    i = i + 1
    file = Dir
Wend

dealYear = Right(Year(Date), 2) 'Retorna 24

getDealNumber = i & "-" & dealYear

End Function

Sub def_export_deal()

Dim dealsFolder As String
Dim pdfName As String
Dim saveLocation As String
Dim ws As Worksheet
Dim dealsLayout As range

dealsFolder = getDealsFolder()
pdfName = "orcamento" & "-" & getDealNumber(dealsFolder)
saveLocation = dealsFolder & pdfName
Set ws = Sheets("layout")
Set dealsLayout = ws.range("B2:J48")

'Exporta o PDF a partir do range definido (dealsLayout)
dealsLayout.ExportAsFixedFormat Type:=xlTypePDF, _
Filename:=saveLocation

End Sub
