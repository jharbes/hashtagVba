Attribute VB_Name = "Module1"
Sub capturar_codigo()

Dim linha As Integer
Dim ultima_linha As Integer
Dim codigo As String

Sheets("Fórmulas de Texto - Parte 1").Activate

ultima_linha = Range("B2").End(xlDown).Row

For linha = 3 To ultima_linha

codigo = Left(Cells(linha, 2).Value, "8")
Cells(linha, 4).Value = codigo


Next

End Sub



Sub capturar_estado()

Dim linha As Integer
Dim ultima_linha As Integer
Dim estado As String

Sheets("Fórmulas de Texto - Parte 2").Activate

ultima_linha = Range("B2").End(xlDown).Row

For linha = 3 To ultima_linha

estado = Right(Cells(linha, 2).Value, "2")
Cells(linha, 4).Value = estado


Next


End Sub
