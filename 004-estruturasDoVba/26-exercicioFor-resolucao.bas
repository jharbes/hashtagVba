Attribute VB_Name = "Module2"
Sub promocao_vendedores()

Dim linha As Integer
Dim media As Double
Dim bim1 As Double
Dim bim2 As Double
Dim valor_necessario As Double


ultima_linha = Range("B10").End(xlDown).Row

For linha = 11 To ultima_linha

    bim1 = Cells(linha, 3)
    bim2 = Cells(linha, 4)
    media = (bim1 + bim2) / 2
    
    If media >= 6000 Then
        Cells(linha, 5).Value = "Promovido"
    Else
        If bim1 >= bim2 Then
            valor_necessario = 12000 - bim1
        Else
            valor_necessario = 12000 - bim2
        End If
        
        Cells(linha, 5).Value = valor_necessario
    End If
        

Next

End Sub
