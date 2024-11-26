Attribute VB_Name = "nd_export_deal"
'Fun��o para obter a pasta de orcamentos
Function getDealsFolder() As String

Dim currentPath As String
Dim dealsFolder As String

' Obt�m o caminho completo do arquivo atual
currentPath = ThisWorkbook.path

' Cria o caminho para a pasta de or�amentos
dealsFolder = currentPath & "\Deals\"

' Cria a pasta de orcamentos se n�o existir
If Dir(dealsFolder, vbDirectory) = "" Then
    MkDir dealsFolder
End If

getDealsFolder = dealsFolder

End Function

'Fun��o para obter o n�mero do orcamento + ano atual
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
