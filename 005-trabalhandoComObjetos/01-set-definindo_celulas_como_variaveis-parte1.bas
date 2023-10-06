Attribute VB_Name = "Module1"
Sub exercicio()

Dim linha As Integer
Dim ultima_linha As Integer

ultima_linha = Range("A2").End(xlDown).Row

For linha = 2 To ultima_linha
    
    Range("D" & linha).Value = 5000
    Range("D" & linha).Font.Bold = True
    Range("D" & linha).Font.Underline = xlUnderlineStyleSingle
    Range("D" & linha).Font.Italic = True
    Range("D" & linha).NumberFormat = _
        "_-[$R$-pt-BR] * #,##0.00_-;-[$R$-pt-BR] * #,##0.00_-;_-[$R$-pt-BR] * ""-""??_-;_-@_-"
    
Next

End Sub
