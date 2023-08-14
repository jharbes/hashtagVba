Attribute VB_Name = "Module1"
Sub repeticao()

Dim linha As Integer
linha = 1

Do Until linha > 10
    
    Cells(linha, 5).Value = "VBA"
    linha = linha + 1

Loop


End Sub




Sub vendas()

Dim linha As Integer
linha = 3

Do Until linha > 10

    Cells(linha, 4).Value = Cells(linha, 3).Value * 0.7
    Cells(linha, 4).NumberFormat = "General"
    linha = linha + 1

Loop

End Sub
