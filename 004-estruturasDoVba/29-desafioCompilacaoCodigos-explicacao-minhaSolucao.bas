Attribute VB_Name = "Module1"
Sub compilar_skus()

Dim marca As String
Dim ano As Integer
Dim sku As String


Sheets("Compilação").Activate

marca = Cells(2, 1).Value
ano = Cells(2, 2).Value

Sheets("Base").Activate

ultima_linha_base = Range("A1").End(xlDown).Row
linha_compilacao = 2

For linha = 2 To ultima_linha_base
    
    Sheets("Base").Activate
    If Cells(linha, 4).Value = ano And Cells(linha, 6) = marca Then
        sku = Cells(linha, 1).Value
        Sheets("Compilação").Activate
        Cells(linha_compilacao, 4).Value = sku
        linha_compilacao = linha_compilacao + 1
    End If
    
    Sheets("Base").Activate
    
Next

Sheets("Compilação").Activate

End Sub
