Attribute VB_Name = "Módulo1"

Sub exemplo_for()

Dim linha As Integer

For linha = 1 To 10

    Range("E" & linha).Value = "VBA"
    'Cells(linha, 5).Value = "VBA"

Next

End Sub


Sub notas()

Dim linha As Integer

For linha = 3 To 11

    If Cells(linha, 3).Value < 5 Then
        Cells(linha, 4).Value = "Reprovado"
    ElseIf Cells(linha, 3).Value < 7 Then
        Cells(linha, 4).Value = "Prova Final"
    Else
        Cells(linha, 4).Value = "Aprovado"
    End If
        
Next

End Sub
