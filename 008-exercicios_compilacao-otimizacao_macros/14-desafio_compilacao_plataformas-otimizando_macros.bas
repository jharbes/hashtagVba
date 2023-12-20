Attribute VB_Name = "Module1"
Sub compila()

Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual


For Each aba In ThisWorkbook.Sheets

    If aba.Name <> "Base" Then
    
        aba.Activate
        Range("B2:H1048576").ClearContents
    
    End If

Next

Sheets("Base").Activate

Dim linha As Integer
Dim linha_registro As Integer
Dim coluna As Integer
Dim mes As String
Dim plataforma As String
Dim volume_extraido As Double


linha = 2

Do Until Cells(linha, 1).Value = ""
    
    mes = Cells(linha, 1).Value
    plataforma = Cells(linha, 3).Value
    volume_extraido = Cells(linha, 4).Value
    
    Sheets(mes).Activate
    
    'Encontra o valor da coluna onde consta a string plataforma
    coluna = Cells.Find(plataforma).Column
    
    linha_registro = Cells(1048576, coluna).End(xlUp).Row + 1
    Cells(linha_registro, coluna).Value = volume_extraido
    
    linha = linha + 1
    
    Sheets("Base").Activate

Loop


Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic


End Sub
