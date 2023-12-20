Attribute VB_Name = "Module1"
Sub compila()

'aumentar a velocidade da macro desligando a visualizacao da macro rodando
'importante desligar apos fim do codigo
Application.ScreenUpdating = False

'aumentar a velocidade da macro desligando o calculo automatico na tabela do excel
'importante desligar apos fim do codigo
Application.Calculation = xlCalculationManual


For Each aba In ThisWorkbook.Sheets

    If aba.Name <> "Base" Then
    
        aba.Range("B2:H1048576").ClearContents
    
    End If

Next


Dim linha As Integer
Dim linha_registro As Integer
Dim coluna As Integer
Dim mes As String
Dim plataforma As String
Dim volume_extraido As Double


linha = 2

Do Until Sheets("Base").Cells(linha, 1).Value = ""
    
    mes = Sheets("Base").Cells(linha, 1).Value
    plataforma = Sheets("Base").Cells(linha, 3).Value
    volume_extraido = Sheets("Base").Cells(linha, 4).Value
    

    'Encontra o valor da coluna onde consta a string plataforma
    coluna = Sheets(mes).Cells.Find(plataforma).Column
    
    linha_registro = Sheets(mes).Cells(1048576, coluna).End(xlUp).Row + 1
    Sheets(mes).Cells(linha_registro, coluna).Value = volume_extraido
    
    linha = linha + 1

Loop


Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic


End Sub
