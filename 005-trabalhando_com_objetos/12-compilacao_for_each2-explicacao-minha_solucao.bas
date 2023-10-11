Attribute VB_Name = "Module1"
Sub compila_funcionarios()

Dim linha As Integer
Dim ultima_linha As Integer
Dim ultima_coluna As Integer
Dim aba As Worksheet


'Apaga toda a lista de funcionario na aba "Resumo Funcionarios"
Sheets("Resumo Funcionarios").Activate
Range("A2:F2").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.ClearContents
Range("A2").Select


'Copia e cola os funcionários de todas as abas para a aba
'"Resumo Funcionarios"
ultima_linha = 2
For Each aba In ThisWorkbook.Sheets

    If aba.Name <> Sheets("Resumo Funcionarios").Name Then
    'ou
    'If aba.Name <> "Resumo Funcionarios" Then
        
        aba.Activate
        Range("A2").Select
        Range(Selection, Selection.End(xlToRight)).Select
        Range(Selection, Selection.End(xlDown)).Select
        Selection.Copy
        Sheets("Resumo Funcionarios").Select
        Range("A" & ultima_linha).Select
        ActiveSheet.Paste
        
        ultima_linha = Range("A2").End(xlDown).Row
        
    End If
        
Next

'Volta o Excel para a aba inicial
Sheets("Resumo Funcionarios").Activate
Range("A2").Select


End Sub
