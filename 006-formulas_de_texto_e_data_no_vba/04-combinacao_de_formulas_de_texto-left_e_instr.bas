Attribute VB_Name = "Module2"
Sub descricao_codigo()

Dim linha As Integer
Dim ultima_linha As Integer
Dim descricao As String
Dim codigo As String
Dim posicao_hifen As Integer


Sheets("Fórmulas de Texto - Parte 4").Activate

ultima_linha = Range("B2").End(xlDown).Row

For linha = 3 To ultima_linha

posicao_hifen = InStr(Cells(linha, 2).Value, "-")
codigo = Left(Cells(linha, 2).Value, posicao_hifen - 2)
descricao = Mid(Cells(linha, 2).Value, posicao_hifen + 2)
Cells(linha, 4).Value = UCase(descricao)
Cells(linha, 5).Value = codigo


Next

End Sub

