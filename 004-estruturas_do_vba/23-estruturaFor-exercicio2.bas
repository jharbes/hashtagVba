Attribute VB_Name = "Module1"
Sub calcula_bonus()

Dim coluna As Integer

'Logica utilizada para calculo da ultima coluna da tabela e
'assim utilizar ela no For
ultima_coluna = Range("B2").End(xlToRight).Column


For coluna = 3 To ultima_coluna

    If Cells(3, coluna).Value < 2500 Then
        Cells(4, coluna).Value = Cells(3, coluna).Value * 0
    ElseIf Cells(3, coluna).Value < 5000 Then
        Cells(4, coluna).Value = Cells(3, coluna).Value * 0.1
    Else
        Cells(4, coluna).Value = Cells(3, coluna).Value * 0.3
    End If
    
Next

End Sub
