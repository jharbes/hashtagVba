Attribute VB_Name = "M�dulo1"
Sub compila_funcionarios()

For Each aba In ThisWorkbook.Sheets

    If aba.Name <> "Resumo Funcionarios" Then
        
    aba.Activate
    
    ult_linha = Range("A1").End(xlDown).Row
    
    Range("A2:F" & ult_linha).Copy
    
    
    End If

Next

End Sub
