Attribute VB_Name = "M�dulo1"

Sub exemplo_for()

Dim linha As Integer

For linha = 1 To 10

    Range("E" & linha).Value = "VBA"
    'Cells(linha, 5).Value = "VBA"

Next

End Sub
