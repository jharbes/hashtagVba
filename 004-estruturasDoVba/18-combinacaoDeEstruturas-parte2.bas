Attribute VB_Name = "Module1"
Sub preenche_prefixo()

Dim linha As Integer
linha = 10

Do While Cells(linha, 4).Value <> ""

    If Cells(linha, 4) = "RJ" Then
        Cells(linha, 5) = "21"
    ElseIf Cells(linha, 4) = "SP" Then
        Cells(linha, 5) = "11"
    ElseIf Cells(linha, 4) = "MG" Then
        Cells(linha, 5) = "31"
    Else
        Cells(linha, 5) = "Desconhecido"
    End If
    
    linha = linha + 1

Loop

End Sub
