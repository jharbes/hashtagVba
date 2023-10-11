Attribute VB_Name = "Módulo1"
Sub compila_funcionarios()

For Each aba In ThisWorkbook.Sheets

    If aba.Name <> "Resumo Funcionarios" Then
        
    aba.Activate
    
    ult_linha = Range("A1").End(xlDown).Row
    
    Range("A2:F" & ult_linha).Copy
    
    Sheets("Resumo Funcionarios").Activate
    
    linha_registro = Range("A100000").End(xlUp).Row + 1
    
    Range("A" & linha_registro).PasteSpecial
    
    
    End If

Next

End Sub
