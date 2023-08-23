Attribute VB_Name = "Module3"
Sub compila()

Dim linha As Integer

ultima_linha = Range("A1").End(xlDown).Row

For linha = 2 To ultima_linha

    If Range("F" & linha).Value = "Antigo" Then
        Rows(linha).Delete
        'Rows(linha & ":" & linha).Delete  'outra opcao para o mesmo comando
    End If

Next


End Sub
