Attribute VB_Name = "Module1"
Sub percorre_celulas()

'observe que a declaracao da variavel celula é opcional
Dim celula As Range

For Each celula In Range("G1:G10")

    celula.Value = "VBA"

Next

End Sub
