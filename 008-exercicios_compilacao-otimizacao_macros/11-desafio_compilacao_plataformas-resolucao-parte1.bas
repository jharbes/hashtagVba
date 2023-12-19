Attribute VB_Name = "Module1"
Sub compila()

Sheets("Base").Activate

Dim linha As Integer
Dim ultima_linha As Integer
Dim mes As String
Dim plataforma As String
Dim volume_extraido As Double


linha = 2

Do Until Cells(linha, 1).Value = ""
    
    mes = Cells(linha, 1).Value
    plataforma = Cells(linha, 3).Value
    volume_extraido = Cells(linha, 4).Value
    
    Sheets(mes).Activate
    
    
    linha = linha + 1

Loop


End Sub
