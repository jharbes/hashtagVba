Attribute VB_Name = "Módulo1"
Sub compila_funcionarios()

Range("A2:F100000").ClearContents

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

Range("A1").Select

End Sub
