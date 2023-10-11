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



Sub vendas2()

Dim linha As Integer
linha = 3

'Essa nova condição fará com que a estrutura seja executada
'ate que ele encontre celulas vazias na coluna 2 e 3
'ao mesmo tempo
Do Until Cells(linha, 2) = "" Or Cells(linha, 3) = ""

    Cells(linha, 4).Value = Cells(linha, 3).Value * 0.7
    Cells(linha, 4).NumberFormat = "General"
    linha = linha + 1

Loop

End Sub


Sub repeticao2()

Dim linha As Integer
linha = 1

Do While linha <= 10
    
    Cells(linha, 5).Value = "VBA"
    linha = linha + 1

Loop

End Sub
