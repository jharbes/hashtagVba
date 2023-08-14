Attribute VB_Name = "Module1"
Sub repeticao()

Dim linha As Integer
linha = 1

Do Until linha > 10

Cells(linha, 5).Value = "VBA"
linha = linha + 1

Loop


End Sub
